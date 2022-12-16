from __future__ import annotations
from dataclasses import dataclass
from ..decorator import enforce
from .. import table_helper as mTb


@enforce.enforce_types
@dataclass(frozen=True)
class RangeValues:
    """
    Range Parts

    .. versionadded:: 0.8.2
    """

    col_start: int
    """Column start such as ``1``"""
    col_end: int
    """Column end such as ``3``"""
    row_start: int
    """Row etart such as ``1``"""
    row_end: int
    """Row end such as ``125``"""

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, RangeValues):
            return False
        return (
            self.col_start == other.col_start
            and self.col_end == other.col_end
            and self.row_start == other.row_start
            and self.row_end == other.row_end
        )

    @staticmethod
    def from_str(range_name: str, zero_index: bool = True) -> RangeValues:
        """
        Gets a ``RangeValues`` instance from a range name.

        Args:
            range_name (str): Range Name such as ``A2:E12``.
            zero_index (bool, optional): Determines if the range name is to be converted to zero index values. Defaults to ``True``.

        Returns:
            RangeValues: Object representing range values.
        """
        return mTb.TableHelper.get_range_values(range_name=str(range_name), zero_index=zero_index)

    def get_range_name(self, zero_index: bool = True) -> str:
        """
        Gets a range name from the current instance.

        Args:
            zero_index (bool, optional): Determines if instance value are treated as zero index values. Defaults to ``True``.

        Returns:
            str: _description_
        """
        start = mTb.TableHelper.make_cell_name(row=self.row_start, col=self.col_start, zero_index=zero_index)
        end = mTb.TableHelper.make_cell_name(row=self.row_end, col=self.col_end, zero_index=zero_index)
        return f"{start}:{end}"
