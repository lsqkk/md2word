"""Mermaid diagram rendering — SVG (preferred) with raster fallback.

SVG approach
━━━━━━━━━━━━━
Mermaid.ink SVG uses <foreignObject> for text, which Word's SVG engine
cannot render. We post-process the SVG to extract text from foreignObjects
and create native SVG <text> elements at correct node-center positions.

Raster fallback
━━━━━━━━━━━━━━━
If SVG rendering fails, falls back to mermaid.ink raster API (JPEG).
"""

from __future__ import annotations

import base64
import re
import subprocess
import tempfile
from pathlib import Path
from typing import NamedTuple


class MermaidResult(NamedTuple):
    """Container for rendered image data with dimensions."""

    image_bytes: bytes
    width: int
    height: int

    @property
    def svg_bytes(self) -> bytes:
        """Deprecated alias — kept for backward compatibility."""
        return self.image_bytes


# ── SVG renderers ───────────────────────────────────────────────────────


def render_via_api_svg(diagram: str) -> MermaidResult | None:
    """Render via mermaid.ink SVG API, then convert foreignObject→text."""
    import requests

    try:
        graph_bytes = diagram.encode("utf-8")
        encoded = base64.urlsafe_b64encode(graph_bytes).decode("ascii")
        resp = requests.get(
            f"https://mermaid.ink/svg/{encoded}", timeout=30,
        )
        resp.raise_for_status()

        patched = _foreignobject_to_text(resp.content)
        w, h = _parse_svg_dimensions(patched)
        return MermaidResult(image_bytes=patched, width=w, height=h)
    except Exception as e:
        print(f"  [WARN] SVG render failed: {e}")
        return None


def render_via_mmdc_svg(diagram: str, mmdc_path: str = "mmdc") -> MermaidResult | None:
    """Render via mmdc CLI to SVG, then convert foreignObject→text."""
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            mmd_file = Path(tmpdir) / "diagram.mmd"
            svg_file = Path(tmpdir) / "diagram.svg"
            mmd_file.write_text(diagram, encoding="utf-8")

            result = subprocess.run(
                [mmdc_path, "-i", str(mmd_file), "-o", str(svg_file)],
                capture_output=True, text=True, timeout=60,
            )
            if result.returncode != 0:
                print(f"  [WARN] mmdc failed: {result.stderr}")
                return None
            if not svg_file.exists():
                return None

            patched = _foreignobject_to_text(svg_file.read_bytes())
            w, h = _parse_svg_dimensions(patched)
            return MermaidResult(image_bytes=patched, width=w, height=h)
    except FileNotFoundError:
        print("  [WARN] mmdc not found. Install: npm install -g @mermaid-js/mermaid-cli")
        return None
    except Exception as e:
        print(f"  [WARN] mmdc SVG render failed: {e}")
        return None


# ── Raster (fallback) renderers ─────────────────────────────────────────


def render_via_api_raster(diagram: str) -> MermaidResult | None:
    """Render via mermaid.ink JPEG API (fallback when SVG fails)."""
    import requests
    from PIL import Image
    from io import BytesIO

    try:
        graph_bytes = diagram.encode("utf-8")
        encoded = base64.urlsafe_b64encode(graph_bytes).decode("ascii")
        resp = requests.get(
            f"https://mermaid.ink/img/{encoded}", timeout=30,
        )
        resp.raise_for_status()
        img = Image.open(BytesIO(resp.content))
        return MermaidResult(image_bytes=resp.content, width=img.width, height=img.height)
    except Exception as e:
        print(f"  [WARN] Raster render failed: {e}")
        return None


def render_via_mmdc_raster(diagram: str, mmdc_path: str = "mmdc") -> MermaidResult | None:
    """Render via mmdc CLI to PNG (fallback)."""
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            mmd_file = Path(tmpdir) / "diagram.mmd"
            png_file = Path(tmpdir) / "diagram.png"
            mmd_file.write_text(diagram, encoding="utf-8")

            result = subprocess.run(
                [mmdc_path, "-i", str(mmd_file), "-o", str(png_file)],
                capture_output=True, text=True, timeout=60,
            )
            if result.returncode != 0:
                return None
            if not png_file.exists():
                return None

            from PIL import Image
            img = Image.open(str(png_file))
            return MermaidResult(image_bytes=png_file.read_bytes(), width=img.width, height=img.height)
    except FileNotFoundError:
        return None
    except Exception as e:
        print(f"  [WARN] mmdc raster render failed: {e}")
        return None


# ── foreignObject→text conversion ───────────────────────────────────────


def _foreignobject_to_text(svg_bytes: bytes) -> bytes:
    """Convert <foreignObject> text to native <text> elements.

    Mermaid renders text inside <foreignObject> (HTML), which Word's SVG
    engine ignores.  We extract the text and create proper SVG <text>
    elements positioned at the node center (the parent <g>'s translate
    already places us at centre — so we use x=0 y=0 with text-anchor=middle).
    """
    svg_str = svg_bytes.decode("utf-8")

    # Regex: match each <g class="node ..." transform="translate(X,Y)">
    # containing a <g class="label"> with a <foreignObject><p>TEXT</p>
    def _replace_label(m: re.Match) -> str:
        label_open = m.group(1)  # <g class="label" transform="...">
        content = m.group(2)     # <foreignObject>...</foreignObject>
        label_close = m.group(3)  # </g>

        p_match = re.search(r"<p[^>]*>(.*?)</p>", content, re.DOTALL)
        if not p_match:
            return m.group(0)

        text = p_match.group(1).strip()
        if not text:
            return label_open + label_close  # remove empty label group

        # Escape XML special chars in text
        text = (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;")
            .replace("'", "&apos;")
        )

        # Extract foreignObject dimensions
        fo_match = re.search(
            r'width="([\d.]+)"\s+height="([\d.]+)"', content
        )
        fo_w = float(fo_match.group(1)) if fo_match else 100
        fo_h = float(fo_match.group(2)) if fo_match else 24

        # Extract label transform offsets (e.g. translate(0, -12))
        t_match = re.search(
            r'transform="translate\(([\d.-]+),\s*([\d.-]+)\)"', label_open
        )
        tx = float(t_match.group(1)) if t_match else 0
        ty = float(t_match.group(2)) if t_match else 0

        # Compute text center: label offset + foreignObject center
        cx = tx + fo_w / 2
        cy = ty + fo_h / 2 + 5  # +5 for baseline offset

        return (
            f'<text x="{cx:.1f}" y="{cy:.1f}" text-anchor="middle"'
            f' dominant-baseline="central" alignment-baseline="central"'
            f' style="text-anchor:middle;font-family:&quot;Microsoft YaHei&quot;,'
            f'SimHei,sans-serif;font-size:16px;fill:#333;">{text}</text>'
        )

    result = re.sub(
        r"(<g\s+class=\"label\"[^>]*>)"
        r"(?:<rect[^>]*/?>)?"
        r"(<foreignObject[^>]*>.*?</foreignObject>)"
        r"(</g>)",
        _replace_label,
        svg_str,
        flags=re.DOTALL,
    )

    return result.encode("utf-8")


# ── Public API ──────────────────────────────────────────────────────────


def render_mermaid(diagram: str, prefer_mmdc: bool = False) -> MermaidResult | None:
    """Render mermaid diagram — SVG preferred, raster fallback."""
    if prefer_mmdc:
        result = render_via_mmdc_svg(diagram)
        if not result:
            result = render_via_mmdc_raster(diagram)
        if not result:
            result = render_via_api_raster(diagram)
    else:
        result = render_via_api_svg(diagram)
        if not result:
            result = render_via_api_raster(diagram)
        if not result:
            result = render_via_mmdc_raster(diagram)

    return result


# ── helpers ─────────────────────────────────────────────────────────────


def _parse_svg_dimensions(svg_bytes: bytes) -> tuple[int, int]:
    """Extract width and height from SVG bytes."""
    from .image_utils import parse_svg_size as _parse_svg

    size = _parse_svg(svg_bytes)
    if size:
        return int(size[0]), int(size[1])
    return 400, 300
