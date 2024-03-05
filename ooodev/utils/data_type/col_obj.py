from __future__ import annotations
from typing import TYPE_CHECKING
from dataclasses import dataclass, field
from weakref import ref
import numbers

from ooodev.utils import table_helper as mTb
from ooodev.utils.validation import check

if TYPE_CHECKING:
    from ooodev.utils.data_type import cell_obj as mCell


@dataclass(frozen=True)
class ColObj:
    """
    Column info.

    .. seealso::
        - :ref:`help_ooodev.utils.data_type.cell_obj.CellObj`

    .. versionadded:: 0.8.2
    """

    # region init

    value: str
    """Column such as ``A``"""
    index: int = field(init=False, repr=False, hash=False)
    """Column Index (zero-based)"""
    cell_obj: "mCell.CellObj | None" = field(repr=False, hash=False, default=None)
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
            AssertionError: if unable to create ``ColObj`` instance.

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
            AssertionError: if unable to create ``ColObj`` instance.

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
        return self.index == other.index if isinstance(other, ColObj) else False

    def __lt__(self, other: object) -> bool:
        try:
            if isinstance(other, str):
                oth = ColObj(other)
                return self.index < oth.index
            i = int(other)  # type: ignore
            return self.index + 1 < i
        except Exception:
            return NotImplemented

    def __le__(self, other: object) -> bool:
        try:
            if isinstance(other, str):
                oth = ColObj(other)
                return self.index <= oth.index
            i = int(other)  # type: ignore
            return self.index + 1 <= i
        except Exception:
            return NotImplemented

    def __gt__(self, other: object) -> bool:
        try:
            if isinstance(other, str):
                oth = ColObj(other)
                return self.index > oth.index
            i = int(other)  # type: ignore
            return self.index + 1 > i
        except Exception:
            return NotImplemented

    def __ge__(self, other: object) -> bool:
        try:
            if isinstance(other, str):
                oth = ColObj(other)
                return self.index >= oth.index
            i = int(other)  # type: ignore
            return self.index + 1 >= i
        except Exception:
            return NotImplemented

    def __add__(self, other: object) -> ColObj:
        if isinstance(other, str):
            try:
                oth = ColObj(other)
                return ColObj.from_int(self.index + oth.index + 2)
            except AssertionError as e:
                raise IndexError from e
            except Exception:
                return NotImplemented
        try:
            if isinstance(other, ColObj):
                return ColObj.from_int(self.index + other.index + 2)
            i = int(other)  # type: ignore
            return ColObj.from_int(self.index + i, True)
        except TypeError:
            # not an int
            return NotImplemented
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            pass
        return NotImplemented

    def __radd__(self, other: object) -> ColObj:
        # angle = sum([col1, col2, col3])
        # will result in TypeError because sum() start with 0
        # this will force a call to __radd__
        return self if other == 0 else self.__add__(other)

    def __sub__(self, other: object) -> ColObj:
        if isinstance(other, str):
            try:
                oth = ColObj(other)
                return ColObj.from_int(self.index - oth.index)
            except AssertionError as e:
                raise IndexError from e
            except Exception:
                return NotImplemented
        try:
            if isinstance(other, ColObj):
                return ColObj.from_int(self.index - other.index)
            i = int(other)  # type: ignore
            return ColObj.from_int(self.index - i, True)
        except TypeError:
            # not an int
            return NotImplemented
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            pass
        return NotImplemented

    def __rsub__(self, other: object) -> ColObj:
        if isinstance(other, str):
            try:
                oth = ColObj(other)
                return ColObj.from_int(oth.index - self.index)
            except AssertionError as e:
                raise IndexError from e
            except Exception:
                return NotImplemented
        try:
            i = int(other)  # type: ignore
            return self.from_int(i - self.index - 1)
        except TypeError:
            # not an int
            return NotImplemented
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            pass
        return NotImplemented

    def __mul__(self, other: object) -> ColObj:
        try:
            i = self.index + 1
            if isinstance(other, str):
                oth = ColObj(other)
                return ColObj.from_int(i * (oth.index + 1))
            if isinstance(other, ColObj):
                return ColObj.from_int(i * (other.index + 1))
            if isinstance(other, numbers.Real):
                return ColObj.from_int(round(i * other))
        except IndexError:
            raise
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            pass
        return NotImplemented

    def __rmul__(self, other: object) -> ColObj:
        return self if other == 0 else self.__mul__(other)

    def __truediv__(self, other: object) -> ColObj:
        try:
            i = self.index + 1
            if isinstance(other, str):
                oth = ColObj(other)
                i2 = oth.index + 1
                check(i >= i2, f"{repr(self)}", "Cannot be divided by lessor number")
                return ColObj.from_int(round(i / i2))
            if isinstance(other, ColObj):
                i2 = other.index + 1
                check(i >= i2, f"{repr(self)}", f"Cannot be divided by lessor {repr(other)}")
                return ColObj.from_int(round(i / i2))
            if isinstance(other, (int, float)):
                check(other != 0, f"{repr(self)}", "Cannot be divided by zero")
                check(i >= other, f"{repr(self)}", "Cannot be divided by lessor number")
            return ColObj.from_int(round(i / other))  # type: ignore
        except IndexError:
            raise
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            pass
        return NotImplemented

    def __rtruediv__(self, other: object) -> ColObj:
        try:
            i = self.index + 1
            if isinstance(other, str):
                oth = ColObj(other)
                return oth.__truediv__(self)
            if isinstance(other, (int, float)):
                check(other != 0, f"{repr(self)}", "Cannot be divided by zero")
                check(other >= i, f"{repr(self)}", "Cannot be divided by greater number")
                return ColObj.from_int(round(other / i))
        except IndexError:
            raise
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            pass
        return NotImplemented

    # endregion dunder methods

    # region properties
    @property
    def next(self) -> ColObj:
        """Gets the next column"""
        # pylint: disable=no-member
        try:
            n = self._next  # type: ignore
            if n() is None:
                raise AttributeError
            return n()
        except AttributeError:
            n = ColObj.from_int(self.index + 2)
            object.__setattr__(self, "_next", ref(n))
            return self._next()  # type: ignore

    @property
    def prev(self) -> ColObj:
        """
        Gets the previous column

        Raises:
            IndexError: If previous column is out of range
        """
        # pylint: disable=no-member
        try:
            p = self._prev  # type: ignore
            if p() is None:
                raise AttributeError
            return p()
        except AttributeError:
            try:
                p = ColObj.from_int(self.index)
                object.__setattr__(self, "_prev", ref(p))
            except AssertionError as e:
                raise IndexError from e
            return self._prev()  # type: ignore

    # endregion properties
