from __future__ import annotations
from dataclasses import dataclass
from .point_positive import PointPositive


@dataclass(frozen=True)
class Offset(PointPositive):
    """Offset x and y values."""

    pass
