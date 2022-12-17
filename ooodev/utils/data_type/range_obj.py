from __future__ import annotations
from dataclasses import dataclass, field
from ..decorator import enforce
from .. import table_helper as mTb
from . import cell_obj as mCo
from . import range_values as mRngValues
from ...office import calc as mCalc

import uno
from ooo.dyn.table.cell_range_address import CellRangeAddress


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
    """Row start such as ``1``"""
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
        if not self.sheet_name:
            try:
                name = mCalc.Calc.get_sheet_name()
                object.__setattr__(self, "sheet_name", name)
            except:
                pass

    @staticmethod
    def from_range(range_val: str | mRngValues.RangeValues | CellRangeAddress) -> RangeObj:
        """
        Gets a ``RangeObj`` from are range name.

        Args:
            range_name (str): Range name such as ``A1:D:34``.

        Returns:
            RangeObj: Object that represents the name range.
        """
        if hasattr(range_val, "typeName") and getattr(range_val, "typeName") == "com.sun.star.table.CellRangeAddress":
            rng = mRngValues.RangeValues.from_range(range_val)
        else:
            rng = range_val

        # return mTb.TableHelper.get_range_obj(range_name=str(range_val))
        if isinstance(rng, mRngValues.RangeValues):
            col_start = mTb.TableHelper.make_column_name(rng.col_start, True)
            col_end = mTb.TableHelper.make_column_name(rng.col_end, True)
            row_start = rng.row_start + 1
            row_end = rng.row_end + 1
            sheet_name = ""

            try:
                if rng.sheet_idx >= 0:
                    sheet = mCalc.Calc.get_sheet(doc=mCalc.Calc.open_doc(), index=rng.sheet_idx)
                    sheet_name = mCalc.Calc.get_sheet_name(sheet)
            except:
                pass
        else:
            parts = mTb.TableHelper.get_range_parts(rng)
            col_start = parts.col_start
            col_end = parts.col_end
            row_start = parts.row_start
            row_end = parts.row_end
            sheet_name = parts.sheet

        return RangeObj(
            col_start=col_start, row_start=row_start, col_end=col_end, row_end=row_end, sheet_name=sheet_name
        )

    def get_range_values(self) -> mRngValues.RangeValues:
        """
        Gets``RangeValues``

        Returns:
            RangeValues: Range Values.
        """
        return mRngValues.RangeValues.from_range(self)

    def to_string(self, include_sheet_name: bool = False) -> str:
        """
        Get a string representation of range

        Args:
            include_sheet_name (bool, optional): If ``True`` and there is a sheet name then it is included in format of ``Sheet1.A2.G3``;
                Otherwise format of ``A2:G3``.Defaults to ``False``.

        Returns:
            str: Gets a string representation such as ``A1:T12``/
        """
        s = f"{self.col_start}{self.row_start}:{self.col_end}{self.row_end}"
        if include_sheet_name and self.sheet_name:
            s = f"{self.sheet_name}." + s
        return s

    def get_cell_range_address(self) -> CellRangeAddress:
        """
        Gets a Cell Range Address

        Returns:
            CellRangeAddress: Cell range Address
        """
        rng = mRngValues.RangeValues.from_range(self)
        return rng.get_cell_range_address()

    def __str__(self) -> str:
        return f"{self.col_start}{self.row_start}:{self.col_end}{self.row_end}"

    def __eq__(self, other: object) -> bool:
        if isinstance(other, RangeObj):
            return self.to_string(True) == other.to_string(True)
        if isinstance(other, mRngValues.RangeValues):
            return str(self) == str(other)
        if isinstance(other, str):
            try:
                oth = RangeObj.from_range(other)
            except Exception:
                return False
            return self.to_string(True) == oth.to_string(True)
        return False
