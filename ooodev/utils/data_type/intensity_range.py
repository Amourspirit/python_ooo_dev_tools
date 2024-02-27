from __future__ import annotations
from dataclasses import dataclass
from ooodev.utils.validation import check
from ooodev.utils.decorator import enforce


@enforce.enforce_types
@dataclass(frozen=True)
class IntensityRange:
    """Represents Intensity Range values. Start and stop values must be from ``0`` to ``100``"""

    start: int
    end: int

    def __post_init__(self) -> None:
        check(
            self.start >= 0 and self.start <= 100,
            f"{self}",
            f"Start value of {self.start} is out of range. Value must be from 0 to 100.",
        )
        check(
            self.end >= 0 and self.end <= 100,
            f"{self}",
            f"End value of {self.end} is out of range. Value must be from 0 to 100.",
        )

    def swap(self) -> IntensityRange:
        """Gets an instance with values swapped."""
        return IntensityRange(self.end, self.start)
