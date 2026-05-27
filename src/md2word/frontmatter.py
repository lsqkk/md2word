"""YAML front matter handling — parse and apply to docx properties."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from docx import Document


def parse_front_matter(text: str) -> tuple[dict[str, Any], str]:
    """Parse YAML front matter (``---…---``) from markdown text.

    Returns (metadata, body) where *metadata* is a dict with keys like
    ``title``, ``author``, ``date``, ``abstract``, ``keywords`` and *body*
    is the remainder of the markdown text.
    """
    metadata: dict[str, Any] = {}

    if not text.startswith("---"):
        return metadata, text

    end = text.find("---", 3)
    if end == -1:
        return metadata, text

    front = text[3:end].strip()
    body = text[end + 3:].lstrip()

    for line in front.split("\n"):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        m = re.match(r'^(\w[\w-]*)\s*:\s*(.*?)\s*$', line)
        if m:
            key = m.group(1).replace("-", "_")
            value = m.group(2).strip().strip("\"'")
            metadata[key] = value

    return metadata, body


def apply_front_matter(doc: Document, metadata: dict[str, Any]) -> None:
    """Apply YAML front matter to docx core properties and add guide paragraphs.

    Maps:
        title   → doc.core_properties.title
        author  → doc.core_properties.author
        date    → doc.core_properties.created
        abstract → rendered as a guide paragraph if ``abstract`` slot exists
        keywords → doc.core_properties.keywords
    """
    cp = doc.core_properties

    if metadata.get("title"):
        cp.title = metadata["title"]

    if metadata.get("author"):
        cp.author = metadata["author"]

    if metadata.get("date"):
        cp.created = _parse_date(metadata["date"])

    if metadata.get("keywords"):
        cp.keywords = metadata["keywords"]

    # abstract is handled by handlers.py which inserts it as a paragraph
    # using the "abstract" style slot


def _parse_date(date_str: str) -> Any:
    """Try to parse a date string into a datetime object."""
    from datetime import datetime

    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except (ValueError, AttributeError):
            continue
    return None
