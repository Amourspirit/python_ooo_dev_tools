from __future__ import annotations
from dataclasses import dataclass
from ooodev.utils.validation import check
from ooodev.utils.data_type.point import Point


@dataclass(frozen=True)
class PointPositive(Point):
    """Represents a X and Y values. Positive Point values. ``x`` and ``y`` values are required to be positive integers"""

    def __post_init__(self) -> None:
        check(
            self.x >= 0 and self.y >= 0,
            f"{self}",
            "Point values must be a positive value",
        )
