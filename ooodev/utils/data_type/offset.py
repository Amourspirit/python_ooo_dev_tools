from __future__ import annotations
from dataclasses import dataclass
from .point_positive import PointPostivie


@dataclass(frozen=True)
class Offset(PointPostivie):
    """Offset x and y values."""

    pass
