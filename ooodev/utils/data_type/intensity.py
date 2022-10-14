from dataclasses import dataclass
from ..validation import check
from ..decorator import enforce


@enforce.enforce_types
@dataclass(frozen=True)
class Intensity:
    """Represents a intensity value from ``0`` to ``100``."""

    Value: int
    """Intensity value."""

    def __post_init__(self) -> None:
        check(
            self.Value >= 0 and self.Value <= 100,
            f"{self}",
            f"Value of {self.Value} is out of range. Value must be from 0 to 100.",
        )
