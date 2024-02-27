from __future__ import annotations
from dataclasses import dataclass
import math
from ooodev.utils.validation import check
from ooodev.utils.decorator import enforce
from ooodev.utils.data_type.base_float_value import BaseFloatValue


@enforce.enforce_types
@dataclass(unsafe_hash=True)
class ImageOffset(BaseFloatValue):
    """Represents a Image Offset value between ``0.0`` and ``1.0``"""

    def __post_init__(self) -> None:
        check(
            self.value >= 0.0 and self.value < 1.0,
            f"{self}",
            f"Value of {self.value} is out of range. Value must be between 0.0 and 1.0",
        )

    def _from_float(self, value: int) -> ImageOffset:
        return ImageOffset(value)

    def __eq__(self, other: object) -> bool:
        try:
            i = float(other)  # type: ignore
            return math.isclose(i, self.value)
        except Exception as e:
            return False
