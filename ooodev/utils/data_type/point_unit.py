from __future__ import annotations
from typing import TYPE_CHECKING
from dataclasses import dataclass

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT


@dataclass(frozen=True)
class PointUnit:
    """Represents a X and Y values."""

    x: UnitT
    y: UnitT
