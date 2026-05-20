"""Mermaid diagram rendering — text → SVG."""

from __future__ import annotations

import base64
import subprocess
import tempfile
from pathlib import Path
from typing import NamedTuple


class MermaidResult(NamedTuple):
    svg_bytes: bytes
    width: int
    height: int


def render_via_api(diagram: str) -> MermaidResult | None:
    """Render mermaid diagram via mermaid.ink API."""
    import requests

    try:
        # Compress and encode the diagram
        graph_bytes = diagram.encode("utf-8")
        encoded = base64.b64encode(graph_bytes).decode("ascii")
        url = f"https://mermaid.ink/svg/{encoded}"

        resp = requests.get(url, timeout=30)
        resp.raise_for_status()

        svg_bytes = resp.content
        w, h = _parse_svg_dimensions(svg_bytes)
        return MermaidResult(svg_bytes=svg_bytes, width=w, height=h)
    except Exception as e:
        print(f"  [WARN] Failed to render mermaid via API: {e}")
        return None


def render_via_mmdc(diagram: str, mmdc_path: str = "mmdc") -> MermaidResult | None:
    """Render mermaid diagram using mmdc (mermaid-cli) locally."""
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

            svg_bytes = svg_file.read_bytes()
            w, h = _parse_svg_dimensions(svg_bytes)
            return MermaidResult(svg_bytes=svg_bytes, width=w, height=h)
    except FileNotFoundError:
        print("  [WARN] mmdc not found. Install: npm install -g @mermaid-js/mermaid-cli")
        return None
    except Exception as e:
        print(f"  [WARN] Failed to render mermaid via mmdc: {e}")
        return None


def render_mermaid(diagram: str, prefer_mmdc: bool = False) -> MermaidResult | None:
    """Render mermaid diagram to SVG.

    Uses mermaid.ink API by default, or mmdc CLI if *prefer_mmdc* is True.
    """
    if prefer_mmdc:
        result = render_via_mmdc(diagram)
        if result:
            return result
        return render_via_api(diagram)
    else:
        result = render_via_api(diagram)
        if result:
            return result
        return render_via_mmdc(diagram)


def _parse_svg_dimensions(svg_bytes: bytes) -> tuple[int, int]:
    """Extract width and height from SVG bytes."""
    from .image_utils import parse_svg_size as _parse_svg

    size = _parse_svg(svg_bytes)
    if size:
        return int(size[0]), int(size[1])
    return 400, 300
