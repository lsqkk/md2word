"""Conversion context and report — replaces global mutable state."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


# ── Severity levels for report entries ─────────────────────────────────────


@dataclass
class ReportEntry:
    """A single report entry with severity."""
    message: str
    severity: str = "warning"  # info / warning / error / critical

    def __str__(self) -> str:
        return f"[{self.severity.upper()}] {self.message}"


# ── Conversion report ──────────────────────────────────────────────────────


@dataclass
class ConversionReport:
    """Tracks warnings and errors during conversion with severity levels.

    Provides structured output that can be queried programmatically.
    """

    entries: list[ReportEntry] = field(default_factory=list)

    # ── convenience accessors ───────────────────────────────────────────

    @property
    def info(self) -> list[str]:
        return [e.message for e in self.entries if e.severity == "info"]

    @property
    def warnings(self) -> list[str]:
        return [e.message for e in self.entries if e.severity == "warning"]

    @property
    def errors(self) -> list[str]:
        return [e.message for e in self.entries if e.severity in ("error", "critical")]

    @property
    def critical(self) -> list[str]:
        return [e.message for e in self.entries if e.severity == "critical"]

    # ── add methods ─────────────────────────────────────────────────────

    def info_msg(self, msg: str) -> None:
        self.entries.append(ReportEntry(message=msg, severity="info"))

    def warn(self, msg: str) -> None:
        self.entries.append(ReportEntry(message=msg, severity="warning"))
        print(f"  ⚠  {msg}")

    def error(self, msg: str) -> None:
        self.entries.append(ReportEntry(message=msg, severity="error"))
        print(f"  ✖  {msg}")

    def add_critical(self, msg: str) -> None:
        self.entries.append(ReportEntry(message=msg, severity="critical"))
        print(f"  🛑  {msg}")

    def has_errors(self) -> bool:
        return any(e.severity in ("error", "critical") for e in self.entries)

    def has_critical(self) -> bool:
        return any(e.severity == "critical" for e in self.entries)

    def summary(self) -> str:
        """Return a structured human-readable summary."""
        lines = []
        info_list = [e for e in self.entries if e.severity == "info"]
        warn_list = [e for e in self.entries if e.severity == "warning"]
        err_list = [e for e in self.entries if e.severity in ("error", "critical")]

        if info_list:
            for e in info_list:
                lines.append(f"  ℹ  {e.message}")
        if warn_list:
            lines.append(f"  ⚠  {len(warn_list)} 个警告")
        if err_list:
            lines.append(f"  ✖  {len(err_list)} 个错误")
        if not err_list and not warn_list:
            lines.append("  ✅ 转换完成，无警告或错误")
        else:
            lines.append("  📋 请查看以上详细信息")
        return "\n".join(lines)


# ── Conversion context (replaces global state) ─────────────────────────────


@dataclass
class ConversionContext:
    """Holds all mutable state during a single convert() call.

    Replaces module-level globals ``_BOOKMARK_COUNTER`` and ``_HEADING_COUNTERS``,
    making conversion re-entrant and safe for batch processing.
    """

    bookmark_counter: int = 0
    heading_counters: list[int] = field(default_factory=lambda: [0] * 7)

    # Stored styles dict
    styles: dict[str, Any] = field(default_factory=dict)

    # Footnote map: placeholder id → Word footnote id
    fn_map: dict[str, int] = field(default_factory=dict)

    # Table counter
    table_counter: int = 0

    # Report
    report: ConversionReport = field(default_factory=ConversionReport)

    # Document instance (set during conversion)
    doc: Any = None

    # ── Bookmark methods ────────────────────────────────────────────────

    def next_bookmark_id(self) -> int:
        self.bookmark_counter += 1
        return self.bookmark_counter

    # ── Heading numbering ───────────────────────────────────────────────

    def reset_heading_counters(self, max_level: int = 6) -> None:
        self.heading_counters = [0] * (max_level + 1)

    def next_heading_number(self, level: int) -> str:
        """Return heading number with offset: h2 => ``1``, h3 => ``1.1``, etc.

        ``h1`` (#) is treated as a document title and never numbered.
        """
        if level <= 1:
            return ""
        dl = level - 1
        self.heading_counters[dl] += 1
        for i in range(dl + 1, len(self.heading_counters)):
            self.heading_counters[i] = 0
        return ".".join(str(self.heading_counters[l]) for l in range(1, dl + 1))
