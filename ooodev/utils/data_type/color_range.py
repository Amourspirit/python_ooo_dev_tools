from __future__ import annotations
from dataclasses import dataclass
from ..validation import check
from ...utils.color import Color


@dataclass(frozen=True)
class ColorRange:
    """Represents color gradient values. Values must be of Type ``Color`` or ``int``."""

    start: Color
    end: Color

    def __post_init__(self) -> None:
        check(
            isinstance(self.start, int) and isinstance(self.end, int),
            f"{self}",
            f"Color range values must be of Type 'Color' or int",
        )
