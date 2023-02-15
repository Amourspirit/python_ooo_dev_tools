from __future__ import annotations
from dataclasses import dataclass
from ..validation import check
from ..decorator import enforce


@enforce.enforce_types
@dataclass(frozen=True)
class Size:
    """Represents a size with postive values."""

    width: int
    height: int

    def __post_init__(self) -> None:
        check(
            self.height >= 0,
            f"{self}",
            f"Value of height ({self.height}) must be a positive number.",
        )
        check(
            self.width >= 0,
            f"{self}",
            f"Value of width ({self.width}) must be a positive number.",
        )

    def swap(self) -> Size:
        """Gets an instance with values swaped."""
        return Size(self.height, self.width)
