from __future__ import annotations
from dataclasses import dataclass, field
from ..decorator import enforce
from .. import table_helper as mTb
from . import cell_obj as mCo


@enforce.enforce_types
@dataclass(frozen=True)
class RangeObj:
    """
    Range Parts

    .. versionadded:: 0.8.2
    """

    col_start: str
    """Column start such as ``A``"""
    col_end: str
    """Column end such as ``C``"""
    row_start: int
    """Row etart such as ``1``"""
    row_end: int
    """Row end such as ``125``"""
    start: mCo.CellObj = field(init=False, repr=False, hash=False)
    """Start Cell Object"""
    end: mCo.CellObj = field(init=False, repr=False, hash=False)
    """End Cell Object"""
    sheet_name: str = ""
    """Sheet name"""

    def __post_init__(self):
        object.__setattr__(self, "col_start", self.col_start.upper())
        object.__setattr__(self, "col_end", self.col_end.upper())

        object.__setattr__(self, "start", mCo.CellObj.from_str(f"{self.col_start}{self.row_start}"))
        object.__setattr__(self, "end", mCo.CellObj.from_str(f"{self.col_end}{self.row_end}"))

    @staticmethod
    def from_str(range_name: str) -> RangeObj:
        """
        Gets a ``RangeObj`` from are range name.

        Args:
            range_name (str): Range name such as ``A1:D:34``.

        Returns:
            RangeObj: Object that represents the name range.
        """
        return mTb.TableHelper.get_range_obj(range_name=range_name)

    def to_string(self, include_sheet_name: bool = False) -> str:
        """
        Get a string representation of range

        Args:
            include_sheet_name (bool, optional): If ``True`` and there is a sheet name then it is included in foramt of ``Sheet1.A2.G3``;
                Otherwider format of ``A2:G3``.Defaults to ``False``.

        Returns:
            str: _description_
        """
        s = f"{self.col_start}{self.row_start}:{self.col_end}{self.row_end}"
        if include_sheet_name and self.sheet_name:
            s = f"{self.sheet_name}." + s
        return s

    def __str__(self) -> str:
        return f"{self.col_start}{self.row_start}:{self.col_end}{self.row_end}"

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, RangeObj):
            return False
        return self.to_string(True) == other.to_string(True)
