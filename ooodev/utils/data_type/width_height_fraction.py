from __future__ import annotations
from dataclasses import dataclass
from ..validation import check
from ..decorator import enforce
import math


# Note that from __future__ import annotations converts annotations to string.
# this means that @enforce.enforce_types will see string as type. This is fine in
# most cases. Especially for built in types.
@enforce.enforce_types
@dataclass(frozen=True)
class WidthHeightFraction:
    """Represents a Width and Height values in decimal values."""

    width: float
    height: float

    def __post_init__(self) -> None:
        check(
            self.width >= 0.0,
            f"{self}",
            f"Width value of {self.width} not valid. Width must be a positive number",
        )
        check(
            self.height >= 0.0,
            f"{self}",
            f"Width value of {self.height} not valid. Width must be a positive number",
        )

    def __eq__(self, other: object) -> bool:
        if isinstance(other, WidthHeightFraction):
            return math.isclose(self.width, other.width) and math.isclose(self.height, other.height)
        return NotImplemented
