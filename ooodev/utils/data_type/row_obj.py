from __future__ import annotations
from dataclasses import dataclass, field
from ..decorator import enforce
from .base_int_value import BaseIntValue
from ..validation import check


@enforce.enforce_types
@dataclass(unsafe_hash=True)
class RowObj(BaseIntValue):
    """
    Column info.

    .. versionadded:: 0.8.2
    """

    index: int = field(init=False, repr=False, hash=False)
    """row Index (zero-based)"""

    def __post_init__(self):
        # must be value of 1 or greater
        check(
            self.value > 0,
            f"{self}",
            f"Expected a value of 1 or greater. Got: {self.value}",
        )
        self.index = self.value - 1

    @staticmethod
    def from_int(num: int, zero_index: bool = False) -> RowObj:
        """
        Gets a ``RowObj`` instance from an interger.

        Args:
            num (int): Row number.
            zero_index (bool, optional): Determines if the row value is treated as zero index. Defaults to ``False``.

        Returns:
            RowObj: Cell Object
        """
        n = num + 1 if zero_index else num
        return RowObj(n)

    def __str__(self) -> str:
        return str(self.value)
