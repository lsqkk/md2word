"""Image handling utilities: download from URL, load local files, fit to page."""

import re
from io import BytesIO
from pathlib import Path

import requests
from docx.shared import Inches, Emu
from PIL import Image as PILImage


# ── SVG helpers ───────────────────────────────────────────────────────────────


def is_svg_source(source: str) -> bool:
    """Quick check if *source* looks like an SVG file (by extension)."""
    return source.lower().endswith(".svg")


def is_svg_data(data: bytes) -> bool:
    """Detect whether *data* contains an SVG image (by content sniffing)."""
    head = data[:512].decode("utf-8", errors="ignore").strip()
    return bool(re.search(r"<svg[\s>]", head, re.IGNORECASE)) or (
        head.startswith("<?xml") and "<svg" in head
    )


def resolve_svg_raw(source: str) -> bytes | None:
    """Load raw SVG bytes from a local path or URL."""
    if source.startswith(("http://", "https://")):
        try:
            resp = requests.get(source, timeout=30)
            resp.raise_for_status()
            data = resp.content
            if not is_svg_data(data):
                print(f"  [WARN] URL does not point to SVG: {source}")
                return None
            return data
        except requests.RequestException as e:
            print(f"  [WARN] Failed to download SVG <{source}>: {e}")
            return None

    path = Path(source)
    if not path.exists():
        print(f"  [WARN] SVG file not found: {source}")
        return None
    return path.read_bytes()


def get_svg_size(
    svg_bytes: bytes, dpi: float = 96
) -> tuple[int, int] | None:
    """Return (width_px, height_px) of SVG at *dpi*.

    Uses resvg when available; falls back to parsing SVG viewBox/attributes.
    """
    try:
        import resvg
    except ImportError:
        return _get_svg_size_fallback(svg_bytes, dpi)

    svg_str = svg_bytes.decode("utf-8", errors="replace")
    try:
        opts = resvg.usvg.Options.default()
        opts.dpi = dpi
        tree = resvg.usvg.Tree.from_str(svg_str, opts)
        w, h = tree.int_size()
        if w == 0 or h == 0:
            print("  [WARN] SVG has zero dimensions — skipping")
            return None
        return (w, h)
    except Exception as e:
        print(f"  [WARN] Failed to parse SVG size: {e}")
        return None


def _get_svg_size_fallback(svg_bytes: bytes, dpi: float = 96) -> tuple[int, int] | None:
    """Parse SVG dimensions from XML attributes (no resvg required)."""
    import xml.etree.ElementTree as ET

    try:
        root = ET.fromstring(svg_bytes)
    except Exception as e:
        print(f"  [WARN] Failed to parse SVG XML: {e}")
        return None

    # Try viewBox first
    vb = root.get("viewBox", "")
    if vb:
        parts = vb.split()
        if len(parts) == 4:
            w_pt = float(parts[2])
            h_pt = float(parts[3])
        else:
            return None
    else:
        # Fall back to width/height attributes (strip units)
        w_str = root.get("width", "0")
        h_str = root.get("height", "0")
        w_pt = float(w_str.rstrip("pt").rstrip("px"))
        h_pt = float(h_str.rstrip("pt").rstrip("px"))

    if w_pt <= 0 or h_pt <= 0:
        return None

    # 1 pt = 1/72 inch, dpi pixels per inch
    scale = dpi / 72.0
    return (max(1, int(w_pt * scale)), max(1, int(h_pt * scale)))


# ── Raster image helpers ──────────────────────────────────────────────────────


def resolve_image(source: str, max_width_inches: float = 5.5) -> BytesIO | None:
    """Load a raster image from a local path or URL, return a normalised stream.

    This function does **not** handle SVG — use ``resolve_svg_raw`` instead.

    Returns:
        A BytesIO stream containing JPEG/PNG data, or None on failure.
    """
    data: bytes | None = None

    if source.startswith(("http://", "https://")):
        try:
            resp = requests.get(source, timeout=30)
            resp.raise_for_status()
            data = resp.content
        except requests.RequestException as e:
            print(f"  [WARN] Failed to download image <{source}>: {e}")
            return None
    else:
        path = Path(source)
        if not path.exists():
            print(f"  [WARN] Local image not found: {source}")
            return None
        try:
            data = path.read_bytes()
        except OSError as e:
            print(f"  [WARN] Failed to read image <{source}>: {e}")
            return None

    if data is None:
        return None

    try:
        img = PILImage.open(BytesIO(data))
        orig_w, orig_h = img.size
        if orig_w > max_width_inches * 96:
            scale = (max_width_inches * 96) / orig_w
            new_w = int(orig_w * scale)
            new_h = int(orig_h * scale)
            img = img.resize((new_w, new_h), PILImage.LANCZOS)

        buf = BytesIO()
        if img.mode in ("RGBA", "P"):
            img.save(buf, format="PNG")
        else:
            if img.mode != "RGB":
                img = img.convert("RGB")
            img.save(buf, format="JPEG", quality=92)
        buf.seek(0)
        return buf
    except Exception as e:
        print(f"  [WARN] Failed to process image <{source}>: {e}")
        return None


def get_image_dimensions(
    image_stream: BytesIO, max_width_inches: float = 5.5
) -> tuple[Inches, Inches]:
    """Return (width, height) in Inches for a given image stream.

    The returned width never exceeds *max_width_inches*; height is scaled
    proportionally.
    """
    img = PILImage.open(image_stream)
    w, h = img.size
    if w > max_width_inches * 96:
        scale = (max_width_inches * 96) / w
        w = int(w * scale)
        h = int(h * scale)
    image_stream.seek(0)
    return Inches(w / 96), Inches(h / 96)
