from dataclasses import dataclass
from ..validation import check
from ..decorator import enforce


@enforce.enforce_types
@dataclass(frozen=True)
class PolySides:
    """Represents Pologon Sides value from ``3`` to ``30``."""

    Value: int
    """Poly Sides value."""

    def __post_init__(self) -> None:
        check(
            self.Value >= 3 and self.Value <= 30,
            f"{self}",
            f"Value of {self.Value} is out of range. Value must be from 3 to 30.",
        )
