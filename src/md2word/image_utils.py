"""Image handling utilities: download from URL, load local files, fit to page."""

import os
import tempfile
from io import BytesIO
from pathlib import Path

import requests
from docx.shared import Inches, Emu
from PIL import Image as PILImage


def resolve_image(source: str, max_width_inches: float = 5.5) -> BytesIO | None:
    """Load an image from a local path or URL, return a BytesIO stream.

    Args:
        source: Local file path or URL starting with http:// or https://.
        max_width_inches: Maximum image width in inches (larger images are
                          proportionally scaled down).

    Returns:
        A BytesIO stream containing the image data, or None on failure.
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
        # Re-encode to a consistent format so python-docx is happy
        img = PILImage.open(BytesIO(data))
        # Preserve original width for size calculation
        orig_w, orig_h = img.size
        if orig_w > max_width_inches * 96:  # 96 DPI screen estimate
            scale = (max_width_inches * 96) / orig_w
            new_w = int(orig_w * scale)
            new_h = int(orig_h * scale)
            img = img.resize((new_w, new_h), PILImage.LANCZOS)

        buf = BytesIO()
        # Convert RGBA/PA to RGB for JPEG, otherwise PNG
        save_format = "PNG"
        if img.mode in ("RGBA", "P"):
            # Keep as PNG to preserve transparency
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
    # Reset stream position for later use
    image_stream.seek(0)
    return Inches(w / 96), Inches(h / 96)
