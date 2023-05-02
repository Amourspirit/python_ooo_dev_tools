from __future__ import annotations
from enum import Enum, IntEnum


class DocType(IntEnum):
    """Document Type"""

    UNKNOWN = 0
    WRITER = 1
    BASE = 2
    CALC = 3
    DRAW = 4
    IMPRESS = 5
    MATH = 6

    def __str__(self) -> str:
        return str(self.value)


class DocTypeStr(str, Enum):
    """Document Type as string"""

    UNKNOWN = "unknown"
    WRITER = "swriter"
    BASE = "sbase"
    CALC = "scalc"
    DRAW = "sdraw"
    IMPRESS = "simpress"
    MATH = "smath"

    def __str__(self) -> str:
        return self.value
