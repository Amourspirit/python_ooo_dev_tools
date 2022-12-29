from __future__ import annotations
from typing import TYPE_CHECKING
from dataclasses import dataclass, field
from weakref import ref
import numbers
from ..validation import check
from .base_int_value import BaseIntValue

if TYPE_CHECKING:
    from . import cell_obj as mCell


@dataclass(unsafe_hash=True)
class RowObj(BaseIntValue):
    """
    Column info.

    .. versionadded:: 0.8.2
    """

    # region init
    index: int = field(init=False, repr=False, hash=False)
    """row Index (zero-based)"""
    cell_obj: "mCell.CellObj | None" = field(repr=False, hash=False, default=None)
    """Cell Object that instance is part of"""

    def __post_init__(self):
        # must be value of 1 or greater
        check(self.value > 0, "RowObj", f"Expected a value of 1 or greater. Got: {self.value}")
        self.index = self.value - 1

    # endregion init

    # region static methods

    @staticmethod
    def from_int(num: int, zero_index: bool = False) -> RowObj:
        """
        Gets a ``RowObj`` instance from an interger.

        Args:
            num (int): Row number.
            zero_index (bool, optional): Determines if the row value is treated as zero index. Defaults to ``False``.

        Raises:
            AssertionError: if unablt to create ``RowObj`` instance.

        Returns:
            RowObj: Cell Object
        """
        if zero_index:
            check(num >= 0, f"{RowObj}", f"Expected a value of 0 or greater. Got: {num}")
        else:
            check(num >= 1, f"{RowObj}", f"Expected a value of 1 or greater. Got: {num}")
        try:
            n = num + 1 if zero_index else num
            return RowObj(n)
        except AssertionError:
            raise
        except Exception as e:
            raise AssertionError from e

    # endregion static methods

    # region duner methods

    def __str__(self) -> str:
        return str(self.value)

    def __int__(self) -> int:
        return self.value

    def __eq__(self, other: object) -> bool:
        if isinstance(other, int):
            return self.value == other
        if not isinstance(other, RowObj):
            return False
        return self.value == other.value

    def __lt__(self, other: object) -> bool:
        try:
            i = int(other)
            return self.value < i
        except Exception:
            return NotImplemented

    def __le__(self, other: object) -> bool:
        try:
            i = int(other)
            return self.value <= i
        except Exception:
            return NotImplemented

    def __gt__(self, other: object) -> bool:
        try:
            i = int(other)
            return self.value > i
        except Exception:
            return NotImplemented

    def __ge__(self, other: object) -> bool:
        try:
            i = int(other)
            return self.value >= i
        except Exception:
            return NotImplemented

    def __add__(self, other: object) -> RowObj:
        if isinstance(other, RowObj):
            return RowObj.from_int(self.value + other.value)
        try:
            i = round(other)
            return RowObj.from_int(self.value + i)
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            pass
        return NotImplemented

    def __radd__(self, other: object) -> RowObj:
        # angle = sum([col1, col2, col3])
        # will result in TypeError becuase sum() start with 0
        # this will force a call to __radd__
        if other == 0:
            return self
        else:
            return self.__add__(other)

    def __sub__(self, other: object) -> RowObj:
        try:
            if isinstance(other, RowObj):
                return RowObj.from_int(self.value - other.value)
            i = round(other)
            return RowObj.from_int(self.value - i)
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            pass
        return NotImplemented

    def __rsub__(self, other: object) -> RowObj:
        try:
            i = round(other)
            return self.from_int(i - self.value)
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            pass
        return NotImplemented

    def __mul__(self, other: object) -> RowObj:
        try:
            if isinstance(other, RowObj):
                return RowObj.from_int(self.value * other.value)
            if isinstance(other, numbers.Real):
                return RowObj.from_int(round(self.value * other))
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            pass
        return NotImplemented

    def __rmul__(self, other: object) -> RowObj:
        if other == 0:
            return self
        else:
            return self.__mul__(other)

    def __truediv__(self, other: object):
        try:
            if isinstance(other, RowObj):
                check(self.value != 0, f"{repr(self)}", f"Cannot be divided by zero")
                return RowObj.from_int(round(self.value / other.value))
            if isinstance(other, numbers.Real):
                check(other != 0, f"{repr(self)}", f"Cannot be divided by zero")
                check(self.value >= other, f"{repr(self)}", f"Cannot be divided by lessor number")
            return RowObj.from_int(round(self.value / other))
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            pass
        return NotImplemented

    def __rtruediv__(self, other: object) -> RowObj:
        try:
            if isinstance(other, numbers.Real):
                check(other != 0, f"{repr(self)}", f"Cannot be divided by zero")
                check(other >= self.value, f"{repr(self)}", f"Cannot be divided by lessor number")
            return RowObj.from_int(round(other / self.value))
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            pass
        return NotImplemented

    # endregion duner methods

    # region properties

    @property
    def next(self) -> RowObj:
        """Gets the nex row"""
        try:
            n = self._next
            if n() is None:
                raise AttributeError
            return n()
        except AttributeError:
            n = RowObj.from_int(self.value + 1)
            object.__setattr__(self, "_next", ref(n))
        return self._next()

    @property
    def prev(self) -> RowObj:
        """
        Gets the prevous row

        Raises:
            IndexError: If previous row is out of range
        """
        try:
            p = self._prev
            if p() is None:
                raise AttributeError
            return p()
        except AttributeError:
            try:
                p = RowObj.from_int(self.value - 1)
                object.__setattr__(self, "_prev", ref(p))
            except AssertionError as e:
                raise IndexError from e
        return self._prev()

    # endregion properties
