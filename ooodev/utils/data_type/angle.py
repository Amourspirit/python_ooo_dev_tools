from dataclasses import dataclass
from ..validation import check
from ..decorator import enforce


@enforce.enforce_types
@dataclass(frozen=True)
class Angle:
    """Represents a angle value from ``0`` to ``360``."""

    Value: int
    """Angle value."""

    def __post_init__(self) -> None:
        check(
            self.Value >= 0 and self.Value <= 360,
            f"{self}",
            f"Value of {self.Value} is out of range. Value must be from 0 to 360.",
        )
