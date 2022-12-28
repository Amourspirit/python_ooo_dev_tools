from __future__ import annotations
from typing import TYPE_CHECKING
from dataclasses import dataclass, field
from . import cell_obj as mCell
from .. import table_helper as mTb

if TYPE_CHECKING:
    try:
        from typing import Self
    except ImportError:
        from typing_extensions import Self


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
    cell_obj: mCell.CellObj | None = field(repr=False, hash=False, default=None)
    """Cell Object that instance is part of"""

    def __post_init__(self):
        object.__setattr__(self, "value", self.value.upper())
        idx = mTb.TableHelper.col_name_to_int(name=self.value, zero_index=True)
        object.__setattr__(self, "index", idx)

    @staticmethod
    def from_str(name: str) -> ColObj:
        """
        Gets a ``ColObj`` instance from a string

        Args:
            name (str): Column letter such as ``A`` or Cell Name such ``A1``

        Returns:
            ColObj: Column Object
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
        return ColObj(mTb.TableHelper.make_column_name(num, zero_index))

    def __int__(self) -> int:
        return self.index

    def __str__(self) -> str:
        return self.value

    def __eq__(self, other: object) -> bool:
        if isinstance(other, str):
            return self.value == other.upper()
        if not isinstance(other, ColObj):
            return False
        return self.index == other.index

    def __lt__(self, other: object) -> bool:
        try:
            if isinstance(other, str):
                oth = ColObj(other)
                return self.index < oth.index
            i = int(other)
            return self.index < i
        except Exception:
            return NotImplemented

    def __le__(self, other: object) -> bool:
        try:
            if isinstance(other, str):
                oth = ColObj(other)
                return self.index <= oth.index
            i = int(other)
            return self.index <= i
        except Exception:
            return NotImplemented

    def __gt__(self, other: object) -> bool:
        try:
            if isinstance(other, str):
                oth = ColObj(other)
                return self.index > oth.index
            i = int(other)
            return self.index > i
        except Exception:
            return NotImplemented

    def __ge__(self, other: object) -> bool:
        try:
            if isinstance(other, str):
                oth = ColObj(other)
                return self.index >= oth.index
            i = int(other)
            return self.index >= i
        except Exception:
            return NotImplemented

    def __add__(self, other: object) -> Self:
        if isinstance(other, str):
            try:
                if isinstance(other, str):
                    oth = ColObj(other)
                    return ColObj.from_int(self.index + oth.index + 2)
            except Exception:
                return NotImplemented
        if isinstance(other, ColObj):
            return ColObj.from_int(self.index + other.index + 2)
        try:
            i = int(other)
            return ColObj.from_int(self.index + i, True)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __radd__(self, other: object) -> Self:
        # angle = sum([col1, col2, col3])
        # will result in TypeError becuase sum() start with 0
        # this will force a call to __radd__
        if other == 0:
            return self
        else:
            return self.__add__(other)

    def __sub__(self, other: object) -> Self:
        if isinstance(other, ColObj):
            return ColObj.from_int(self.index - other.index)
        if isinstance(other, str):
            try:
                oth = ColObj(other)
                return ColObj.from_int(self.index - oth.index)
            except Exception:
                raise NotImplemented
        try:
            i = int(other)
            return ColObj.from_int(self.index - i, True)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented

    def __rsub__(self, other: object) -> Self:
        if isinstance(other, str):
            try:
                oth = ColObj(other)
                return ColObj.from_int(oth.index - self.index)
            except Exception:
                raise NotImplemented
        try:
            i = int(other)
            return self.from_int(i - self.index - 1)
        except AssertionError:
            raise
        except Exception:
            return NotImplemented
