from __future__ import annotations
import string
from dataclasses import dataclass, field
from . import col_obj as mCol
from . import range_obj as mRngObj
from . import row_obj as mRow
from .. import table_helper as mTb
from ...office import calc as mCalc

from ooo.dyn.table.cell_address import CellAddress


@dataclass(frozen=True)
class CellObj:
    """
    Cell Parts

    .. versionadded:: 0.8.2
    """

    col: str
    """Column such as ``A``"""
    row: int
    """Row such as ``125``"""
    sheet_idx: int = -1
    """Sheet index that this cell value belongs to"""
    range_obj: mRngObj.RangeObj | None = field(repr=False, hash=False, default=None)
    """Range Object that instance is part of"""

    def __post_init__(self):
        object.__setattr__(self, "col", self.col.upper())
        if self.sheet_idx < 0:
            if self.range_obj:
                if self.range_obj.sheet_idx >= 0:
                    object.__setattr__(self, "sheet_idx", self.range_obj.sheet_idx)
            else:
                try:
                    idx = mCalc.Calc.get_sheet_index()
                    object.__setattr__(self, "sheet_idx", idx)
                except:
                    pass

    @staticmethod
    def from_cell(cell_val: str | CellAddress) -> CellObj:
        """
        Gets a ``CellObj`` instance from a string

        Args:
            cell_val (str | CellAddress): Cell Name such ``A23`` or ``Sheet1.A23``, or Cell Address object.

        Returns:
            CellObj: Cell Object

        Note:
            If a range name such as ``A23:G45`` or ``Sheet1.A23:G45`` then only the first cell is used.
        """
        if isinstance(cell_val, str):
            # split will cover if a range is passed in, return first cell
            parts = mTb.TableHelper.get_cell_parts(cell_val)
            idx = -1
            if parts.sheet:
                try:
                    sheet = mCalc.Calc.get_sheet(doc=mCalc.Calc.open_doc(), sheet_name=parts.sheet)
                    idx = mCalc.Calc.get_sheet_index(sheet=sheet)
                except:
                    pass
            return CellObj(col=parts.col, row=parts.row, sheet_idx=idx)

        # CellAddress
        col = mTb.TableHelper.make_column_name(cell_val.Column, True)
        return CellObj(col=col, row=cell_val.Row + 1, sheet_idx=cell_val.Sheet)

    @staticmethod
    def from_idx(col_idx: int, row_idx: int, sheet_idx: int = -1) -> CellObj:
        """
        Gets a ``CellObj`` from zero-based col and row indexes

        Args:
            col_idx (int): Column index
            row_idx (int): Row index
            sheet_idx (int, optional): Sheet index

        Returns:
            CellObj: Cell object
        """
        col = mTb.TableHelper.make_column_name(col=col_idx, zero_index=True)
        return CellObj(col=col, row=row_idx + 1, sheet_idx=sheet_idx)

    def get_cell_address(self) -> CellAddress:
        """
        Gets a cell address

        Returns:
            CellAddress: Cell Address
        """
        column = mTb.TableHelper.col_name_to_int(self.col, True)
        return CellAddress(Sheet=self.sheet_idx, Column=column, Row=self.row - 1)

    def get_range_obj(self) -> mRngObj.RangeObj:
        """
        Gets a Range object that has start and end column set to this instance cell values.

        Returns:
            mRngObj.RangeObj: Range Object
        """
        return mRngObj.RangeObj(
            col_start=self.col, col_end=self.col, row_start=self.row, row_end=self.row, sheet_idx=self.sheet_idx
        )

    def __str__(self) -> str:
        return f"{self.col}{self.row}"

    def __eq__(self, other: object) -> bool:
        if isinstance(other, CellObj):
            return self.sheet_idx == other.sheet_idx and self.col == other.col and self.row == other.row
        if isinstance(other, str):
            return str(self) == other.upper()
        return False

    @property
    def col_info(self) -> mCol.ColObj:
        """Gets Column Info"""
        try:
            return self._col_info
        except AttributeError:
            object.__setattr__(self, "_col_info", mCol.ColObj(value=self.col, cell_obj=self))
        return self._col_info

    @property
    def row_info(self) -> mRow.RowObj:
        """Gets Row Info"""
        try:
            return self._row_info
        except AttributeError:
            object.__setattr__(self, "_row_info", mRow.RowObj(value=self.row, cell_obj=self))
        return self._row_info
