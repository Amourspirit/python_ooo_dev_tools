from __future__ import annotations
from dataclasses import dataclass
from typing import overload
from ..decorator import enforce
from .. import table_helper as mTb
from . import range_obj as mRngObj


@enforce.enforce_types
@dataclass(frozen=True)
class RangeValues:
    """
    Range Parts. Intended to be zero-based indexs

    .. versionadded:: 0.8.2
    """

    col_start: int
    """Column start such as ``0``"""
    col_end: int
    """Column end such as ``3``"""
    row_start: int
    """Row start such as ``0``"""
    row_end: int
    """Row end such as ``125``"""

    def __eq__(self, other: object) -> bool:
        if isinstance(other, RangeValues):
            return (
                self.col_start == other.col_start
                and self.col_end == other.col_end
                and self.row_start == other.row_start
                and self.row_end == other.row_end
            )
        if isinstance(other, mRngObj.RangeObj):
            return str(self) == str(other)
        if isinstance(other, str):
            return str(self) == other.upper()
        return False

    def __str__(self) -> str:
        start = mTb.TableHelper.make_cell_name(row=self.row_start, col=self.col_start, zero_index=True)
        end = mTb.TableHelper.make_cell_name(row=self.row_end, col=self.col_end, zero_index=True)
        return f"{start}:{end}"

    # region from_range()
    @overload
    @staticmethod
    def from_range(range_val: mRngObj.RangeObj) -> RangeValues:
        ...

    @overload
    @staticmethod
    def from_range(range_val: str) -> RangeValues:
        ...

    @staticmethod
    def from_range(range_val: str | mRngObj.RangeObj) -> RangeValues:
        """
        Gets a ``RangeValues`` instance from a range

        Args:
            range (str | mRngObj.RangeObj): Range as object or string (``A2:G23``).

        Returns:
            RangeValues: Object representing range values.
        """
        return mTb.TableHelper.get_range_values(range_name=str(range_val))

    # endregion from_range()

    def get_range_obj(self) -> mRngObj.RangeObj:
        """
        Gets a ``RangeObj``

        Returns:
            mRngObj.RangeObj: Range object.
        """
        return mRngObj.RangeObj.from_range(str(self))
