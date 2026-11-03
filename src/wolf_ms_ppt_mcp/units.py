"""Unit conversions between inches (agent-facing) and EMU (python-pptx)."""

from __future__ import annotations

from pptx.util import Emu

EMU_PER_INCH = 914_400


def inches_to_emu(value: float) -> Emu:
    """Convert inches to EMU (English Metric Units)."""
    return Emu(int(round(value * EMU_PER_INCH)))


def emu_to_inches(value: int | Emu) -> float:
    """Convert EMU to inches, rounded to 3 decimals."""
    return round(int(value) / EMU_PER_INCH, 3)
