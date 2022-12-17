from __future__ import annotations
from dataclasses import dataclass, field
from ..decorator import enforce
from .. import table_helper as mTb


@enforce.enforce_types
@dataclass(frozen=True)
class ColObj:
    """
    Column info.

    .. versionadded:: 0.8.2
    """

    value: str
    """Column such as ``A``"""
    index: int = field(init=False, repr=False, hash=False)
    """Column Index (zero-based)"""

    def __post_init__(self):
        object.__setattr__(self, "value", self.value.upper())
        idx = mTb.TableHelper.col_name_to_int(name=self.value, zero_index=True)
        object.__setattr__(self, "index", idx)

    @staticmethod
    def from_str(name: str) -> ColObj:
        """
        Gets a ``ColObj`` instance from a string

        Args:
            name (str): Cell Name such ``A1``

        Returns:
            CellObj: Cell Object
        """
        num = mTb.TableHelper.col_name_to_int(name=name)
        return ColObj(mTb.TableHelper.make_column_name(num))

    @staticmethod
    def from_int(num: int, zero_index: bool = False) -> ColObj:
        """
        Gets a ``ColObj`` instance from an integer.

        Args:
            num (int): Column number.
            zero_index (bool, optional): Determines if the column number is treated as zero index. Defaults to ``False``.

        Returns:
            ColObj: Cell Object
        """
        return ColObj(mTb.TableHelper.make_column_name(num=num, zero_index=zero_index))

    def __int__(self) -> int:
        return self.index

    def __str__(self) -> str:
        return self.value

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, ColObj):
            return False
        return self.index == other.index

    def __lt__(self, other: object) -> bool:
        try:
            i = int(other)
            return self.index < i
        except Exception:
            return NotImplemented

    def __le__(self, other: object) -> bool:
        try:
            i = int(other)
            return self.index <= i
        except Exception:
            return NotImplemented

    def __gt__(self, other: object) -> bool:
        try:
            i = int(other)
            return self.index > i
        except Exception:
            return NotImplemented

    def __ge__(self, other: object) -> bool:
        try:
            i = int(other)
            return self.index >= i
        except Exception:
            return NotImplemented
