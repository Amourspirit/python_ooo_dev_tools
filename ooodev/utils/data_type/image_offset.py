from dataclasses import dataclass
from ..validation import check
from ..decorator import enforce


@enforce.enforce_types
@dataclass(frozen=True)
class ImageOffset:
    """Represents a Image Offset value between ``0.0`` and ``1.0``"""

    Value: float
    """Offset value."""

    def __post_init__(self) -> None:
        check(
            (self.Value < 0.0 or self.Value >= 1.0) == False,
            f"{self}",
            f"Value of {self.Value} is out of range. Value must be between 0.0 and 1.0",
        )
