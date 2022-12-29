from __future__ import annotations
from typing import TYPE_CHECKING
from dataclasses import dataclass, field
from weakref import ref
from . import cell_obj as mCell
from .. import table_helper as mTb
from ..validation import check

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

    # region init

    value: str
    """Column such as ``A``"""
    index: int = field(init=False, repr=False, hash=False)
    """Column Index (zero-based)"""
    cell_obj: mCell.CellObj | None = field(repr=False, hash=False, default=None)
    """Cell Object that instance is part of"""

    def __post_init__(self):

        object.__setattr__(self, "value", self.value.upper())
        try:
            idx = mTb.TableHelper.col_name_to_int(name=self.value, zero_index=True)
        except ValueError as e:
            raise AssertionError from e
        check(idx >= 0, f"{self}", f"Expected a value index of 0 or greater. Got: {idx}")
        object.__setattr__(self, "index", idx)

    # endregion init

    # region static methods

    @staticmethod
    def from_str(name: str) -> ColObj:
        """
        Gets a ``ColObj`` instance from a string

        Args:
            name (str): Column letter such as ``A`` or Cell Name such ``A1``

        Raises:
            AssertionError: if unablt to create ``ColObj`` instance.

        Returns:
            ColObj: Column Object
        """
        try:
            num = mTb.TableHelper.col_name_to_int(name=name)
            return ColObj(mTb.TableHelper.make_column_name(num))
        except AssertionError:
            raise
        except Exception as e:
            raise AssertionError from e

    @staticmethod
    def from_int(num: int, zero_index: bool = False) -> ColObj:
        """
        Gets a ``ColObj`` instance from an integer.

        Args:
            num (int): Column number.
            zero_index (bool, optional): Determines if the column number is treated as zero index. Defaults to ``False``.

        Raises:
            AssertionError: if unablt to create ``ColObj`` instance.

        Returns:
            ColObj: Cell Object
        """
        if zero_index:
            check(num >= 0, f"{ColObj}", f"Expected a value of 0 or greater. Got: {num}")
        else:
            check(num >= 1, f"{ColObj}", f"Expected a value of 1 or greater. Got: {num}")
        try:
            col_name = mTb.TableHelper.make_column_name(num, zero_index)
            return ColObj(col_name)
        except AssertionError:
            raise
        except Exception as e:
            raise AssertionError from e

    # endregion static methods

    # region dunder methods

    def __int__(self) -> int:
        return self.index + 1

    def __str__(self) -> str:
        return self.value

    def __eq__(self, other: object) -> bool:
        if isinstance(other, str):
            return self.value == other.upper()
        if isinstance(other, int):
            return self.index + 1 == other
        if not isinstance(other, ColObj):
            return False
        return self.index == other.index

    def __lt__(self, other: object) -> bool:
        try:
            if isinstance(other, str):
                oth = ColObj(other)
                return self.index < oth.index
            i = int(other)
            return self.index + 1 < i
        except Exception:
            return NotImplemented

    def __le__(self, other: object) -> bool:
        try:
            if isinstance(other, str):
                oth = ColObj(other)
                return self.index <= oth.index
            i = int(other)
            return self.index + 1 <= i
        except Exception:
            return NotImplemented

    def __gt__(self, other: object) -> bool:
        try:
            if isinstance(other, str):
                oth = ColObj(other)
                return self.index > oth.index
            i = int(other)
            return self.index + 1 > i
        except Exception:
            return NotImplemented

    def __ge__(self, other: object) -> bool:
        try:
            if isinstance(other, str):
                oth = ColObj(other)
                return self.index >= oth.index
            i = int(other)
            return self.index + 1 >= i
        except Exception:
            return NotImplemented

    def __add__(self, other: object) -> Self:
        if isinstance(other, str):
            try:
                if isinstance(other, str):
                    oth = ColObj(other)
                    return ColObj.from_int(self.index + oth.index + 2)
            except AssertionError as e:
                raise IndexError from e
            except Exception:
                return NotImplemented
        try:
            if isinstance(other, ColObj):
                return ColObj.from_int(self.index + other.index + 2)
            i = int(other)
            return ColObj.from_int(self.index + i, True)
        except AssertionError as e:
            raise IndexError from e
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
        if isinstance(other, str):
            try:
                oth = ColObj(other)
                return ColObj.from_int(self.index - oth.index)
            except AssertionError as e:
                raise IndexError from e
            except Exception:
                raise NotImplemented
        try:
            if isinstance(other, ColObj):
                return ColObj.from_int(self.index - other.index)
            i = int(other)
            return ColObj.from_int(self.index - i, True)
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            return NotImplemented

    def __rsub__(self, other: object) -> Self:
        if isinstance(other, str):
            try:
                oth = ColObj(other)
                return ColObj.from_int(oth.index - self.index)
            except AssertionError as e:
                raise IndexError from e
            except Exception:
                raise NotImplemented
        try:
            i = int(other)
            return self.from_int(i - self.index - 1)
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            return NotImplemented

    # endregion dunder methods

    # region properties
    @property
    def next(self) -> ColObj:
        """Gets the next column"""
        try:
            n = self._next
            if n() is None:
                raise AttributeError
            return n()
        except AttributeError:
            n = ColObj.from_int(self.index + 2)
            object.__setattr__(self, "_next", ref(n))
            return self._next()

    @property
    def prev(self) -> ColObj:
        """
        Gets the previous column

        Raises:
            IndexError: If prevous column is out of range
        """
        try:
            p = self._prev
            if p() is None:
                raise AttributeError
            return p()
        except AttributeError:
            try:
                p = ColObj.from_int(self.index)
                object.__setattr__(self, "_prev", ref(p))
            except AssertionError as e:
                raise IndexError from e
            return self._prev()

    # endregion properties
