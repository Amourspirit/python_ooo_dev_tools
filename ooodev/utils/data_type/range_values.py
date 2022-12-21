from __future__ import annotations
from typing import TYPE_CHECKING
from dataclasses import dataclass
from typing import cast, overload

from . import cell_obj as mCellObj
from . import cell_values as mCellVals
from . import range_obj as mRngObj
from .. import table_helper as mTb
from ...office import calc as mCalc
from ..decorator import enforce

import uno
from ooo.dyn.table.cell_range_address import CellRangeAddress

if TYPE_CHECKING:
    from ooo.lo.table.cell_address import CellAddress


@enforce.enforce_types
@dataclass(frozen=True)
class RangeValues:
    """
    Range Parts. Intended to be zero-based indexes

    .. versionadded:: 0.8.2
    """

    col_start: int
    """Column start such as ``0``. Must be non-negative integer"""
    col_end: int
    """Column end such as ``3``. Must be non-negative integer"""
    row_start: int
    """Row start such as ``0``. Must be non-negative integer"""
    row_end: int
    """Row end such as ``125``. Must be non-negative integer"""
    sheet_idx: int = -1
    """Sheet index that this range value belongs to"""

    def __post_init__(self):
        cr_vals = (self.col_start, self.row_start, self.col_end, self.row_end)
        for val in cr_vals:
            if val < 0:
                raise ValueError(
                    f"All indexes must be greater than 0. Column Range ({self.col_start}:{self.col_end}), Row Range - ({self.row_start}:{self.row_end})"
                )

        if self.sheet_idx < 0:
            try:
                idx = mCalc.Calc.get_sheet_index()
                object.__setattr__(self, "sheet_idx", idx)
            except:
                pass

    def __eq__(self, other: object) -> bool:
        if isinstance(other, RangeValues):
            return (
                self.col_start == other.col_start
                and self.col_end == other.col_end
                and self.row_start == other.row_start
                and self.row_end == other.row_end
                and self.sheet_idx == other.sheet_idx
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
    def from_range(range_val: CellRangeAddress) -> RangeValues:
        ...

    @overload
    @staticmethod
    def from_range(range_val: mRngObj.RangeObj) -> RangeValues:
        ...

    @overload
    @staticmethod
    def from_range(range_val: str) -> RangeValues:
        ...

    @staticmethod
    def from_range(range_val: str | mRngObj.RangeObj | CellRangeAddress) -> RangeValues:
        """
        Gets a ``RangeValues`` instance from a range

        Args:
            range (str | mRngObj.RangeObj | CellRangeAddress): Range as object or string (``A2:G23``).

        Returns:
            RangeValues: Object representing range values.
        """
        if isinstance(range_val, mRngObj.RangeObj):
            col_start = mTb.TableHelper.col_name_to_int(range_val.col_start, True)
            col_end = mTb.TableHelper.col_name_to_int(range_val.col_end, True)
            row_start = range_val.row_start - 1
            row_end = range_val.row_end - 1
            sheet_idx = range_val.sheet_idx
        elif isinstance(range_val, str):
            parts = mTb.TableHelper.get_range_parts(range_val)
            col_start = mTb.TableHelper.col_name_to_int(parts.col_start, True)
            col_end = mTb.TableHelper.col_name_to_int(parts.col_end, True)
            row_start = parts.row_start - 1
            row_end = parts.row_end - 1
            sheet_name = parts.sheet
            sheet_idx = -1
            if sheet_name:
                sheet = mCalc.Calc.get_sheet(doc=mCalc.Calc.get_current_doc(), sheet_name=sheet_name)
                sheet_idx = mCalc.Calc.get_sheet_index(sheet)
        else:
            # CellRange
            col_start = range_val.StartColumn
            col_end = range_val.EndColumn
            row_start = range_val.StartRow
            row_end = range_val.EndRow
            sheet_idx = range_val.Sheet

        return RangeValues(
            col_start=col_start, row_start=row_start, col_end=col_end, row_end=row_end, sheet_idx=sheet_idx
        )

    # endregion from_range()

    def get_range_obj(self) -> mRngObj.RangeObj:
        """
        Gets a ``RangeObj``

        Returns:
            RangeObj: Range object.
        """
        return mRngObj.RangeObj.from_range(self)

    def get_cell_range_address(self) -> CellRangeAddress:
        """
        Gets a Cell Range Address

        Returns:
            CellRangeAddress: Cell range Address
        """
        return CellRangeAddress(
            Sheet=self.sheet_idx,
            StartColumn=self.col_start,
            StartRow=self.row_start,
            EndColumn=self.col_end,
            EndRow=self.row_end,
        )

    def is_single_col(self) -> bool:
        """
        Gets if instance is a single column or multi-column

        Returns:
            bool: ``True`` if single column; Otherwise, ``False``

        Note:
            If instance is a single cell address then ``True`` is returned.
        """
        return self.col_start == self.col_end

    def is_single_row(self) -> bool:
        """
        Gets if instance is a single row or multi-row

        Returns:
            bool: ``True`` if single row; Otherwise, ``False``

        Note:
            If instance is a single cell address then ``True`` is returned.
        """
        return self.row_start == self.row_end

    def is_single_cell(self) -> bool:
        """
        Gets if a instance is a single cell or a range

        Returns:
            bool: ``True`` if single cell; Otherwise, ``False``
        """
        return self.is_single_col() and self.is_single_row()

    # region contains()

    @overload
    def contains(self, cell_obj: mCellObj.CellObj) -> bool:
        ...

    @overload
    def contains(self, cell_addr: CellAddress) -> bool:
        ...

    @overload
    def contains(self, cell_vals: mCellVals.CellValues) -> bool:
        ...

    @overload
    def contains(self, cell_name: str) -> bool:
        ...

    def contains(self, *args, **kwargs) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            cell_obj (CellObj): Cell object
            cell_addr (CellAddress): Cell address
            cell_vals (CellValues): Cell Values
            cell_name (str): Cell name

        Returns:
            bool: ``True`` if instance contains cell; Otherwishe, ``False``.

        Note:
            If cell input contains sheet info the it is use in comparsion.
            Otherwise sheet is ignored.
        """
        ordered_keys = (1,)
        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        def get_kwargs() -> dict:
            ka = {}
            if kargs_len == 0:
                return ka
            valid_keys = ("cell_obj", "cell_vals", "cell_name", "cell_addr")
            check = all(key in valid_keys for key in kwargs.keys())
            if not check:
                raise TypeError("contains() got an unexpected keyword argument")
            for key in valid_keys:
                if key in kwargs:
                    ka[1] = kwargs[key]
                    break
            return ka

        if count != 1:
            raise TypeError("contains() got an invalid number of arguments")

        kargs = get_kwargs()
        for i, arg in enumerate(args):
            kargs[ordered_keys[i]] = arg

        arg = kargs[1]
        if isinstance(arg, mCellObj.CellObj):
            cvs = arg.get_cell_values()
            col = cvs.col
            row = cvs.row
            idx = cvs.sheet_idx
        elif isinstance(arg, mCellVals.CellValues):
            col = arg.col
            row = arg.row
            idx = arg.sheet_idx
        elif isinstance(arg, str):
            cvs = mCellVals.CellValues.from_cell(arg)
            col = cvs.col
            row = cvs.row
            idx = cvs.sheet_idx
        else:
            # CellAddress
            ca = cast("CellAddress", arg)
            col = ca.Column
            row = ca.Row
            idx = ca.Sheet

        contains = True
        if idx >= 0:
            contains = contains and self.sheet_idx == idx
        contains = contains and self.col_start <= col
        contains = contains and self.row_start <= row
        contains = contains and self.col_end >= col
        contains = contains and self.row_end >= row
        return contains

    # endregion contains()
