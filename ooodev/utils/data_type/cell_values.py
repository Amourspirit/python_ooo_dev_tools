from __future__ import annotations
from dataclasses import dataclass
from ..decorator import enforce


@enforce.enforce_types
@dataclass(frozen=True)
class CellValues:
    """
    Cell Parts

    .. versionadded:: 0.8.2
    """

    col: int
    """Column such as ``1``"""
    row: int
    """Row such as ``125``"""

    def __str__(self) -> str:
        return f"{self.col}{self.row}"
