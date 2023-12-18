from __future__ import annotations
from dataclasses import dataclass
from ooodev.units import UnitT


@dataclass(frozen=True)
class PointUnit:
    """Represents a X and Y values."""

    x: UnitT
    y: UnitT
