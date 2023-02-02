from __future__ import annotations
from typing import TYPE_CHECKING
from dataclasses import dataclass
import math
from ..validation import check
from ..decorator import enforce
from .base_float_value import BaseFloatValue


if TYPE_CHECKING:
    try:
        from typing import Self
    except ImportError:
        from typing_extensions import Self


@enforce.enforce_types
@dataclass(unsafe_hash=True)
class ImageOffset(BaseFloatValue):
    """Represents a Image Offset value between ``0.0`` and ``1.0``"""

    def __post_init__(self) -> None:
        check(
            (self.value < 0.0 or self.value >= 1.0) == False,
            f"{self}",
            f"Value of {self.value} is out of range. Value must be between 0.0 and 1.0",
        )

    def _from_float(self, value: int) -> Self:
        return ImageOffset(value)

    def __eq__(self, other: object) -> bool:

        try:
            i = float(other)
            return math.isclose(i, self.value)
        except Exception as e:
            return False
