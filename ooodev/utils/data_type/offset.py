from __future__ import annotations
from dataclasses import dataclass
from ooodev.utils.data_type.point_positive import PointPositive


@dataclass(frozen=True)
class Offset(PointPositive):
    """Offset x and y values."""

    pass
