from __future__ import annotations
from typing import TYPE_CHECKING, cast
from dataclasses import dataclass, field
from typing import overload
from weakref import ref
from . import cell_values as mCellVals
from . import col_obj as mCol
from . import range_obj as mRngObj
from . import row_obj as mRow
from .. import table_helper as mTb
from ...office import calc as mCalc
from ..validation import check

from ooo.dyn.table.cell_address import CellAddress

if TYPE_CHECKING:
    try:
        from typing import Self
    except ImportError:
        from typing_extensions import Self


@dataclass(frozen=True)
class CellObj:
    """
    Cell Parts

    .. versionadded:: 0.8.2
    """

    # region init

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
        try:
            # convert col to index for the purpose of validataion
            _ = mTb.TableHelper.col_name_to_int(name=self.col)
        except ValueError as e:
            raise AssertionError from e
        check(self.row >= 1, f"{self}", f"Expected a row of 1 or greater. Got: {self.row}")
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

    # endregion init

    # region static methods

    # region from_cell()

    @overload
    @staticmethod
    def from_cell(cell_val: str) -> CellObj:
        ...

    @overload
    @staticmethod
    def from_cell(cell_val: CellObj) -> CellObj:
        ...

    @overload
    @staticmethod
    def from_cell(cell_val: CellAddress) -> CellObj:
        ...

    @overload
    @staticmethod
    def from_cell(cell_val: mCellVals.CellValues) -> CellObj:
        ...

    @staticmethod
    def from_cell(cell_val: str | CellAddress | mCellVals.CellValues | CellObj) -> CellObj:
        """
        Gets a ``CellObj`` instance from a string

        Args:
            cell_val (str | CellAddress | CellValues | CellObj): Cell value. If ``cell_val`` is `CellObj`` then that instance is returned.

        Returns:
            CellObj: Cell Object

        Note:
            If a range name such as ``A23:G45`` or ``Sheet1.A23:G45`` then only the first cell is used.
        """
        if isinstance(cell_val, CellObj):
            return cell_val

        if isinstance(cell_val, str):
            # split will cover if a range is passed in, return first cell
            parts = mTb.TableHelper.get_cell_parts(cell_val)
            idx = -1
            if parts.sheet:
                try:
                    sheet = mCalc.Calc.get_sheet(doc=mCalc.Calc.get_current_doc(), sheet_name=parts.sheet)
                    idx = mCalc.Calc.get_sheet_index(sheet=sheet)
                except:
                    pass
            return CellObj(col=parts.col, row=parts.row, sheet_idx=idx)

        cv = mCellVals.CellValues.from_cell(cell_val)
        col = mTb.TableHelper.make_column_name(cv.col, True)
        return CellObj(col=col, row=cv.row + 1, sheet_idx=cv.sheet_idx)

    # endregion from_cell()

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

    # endregion static methods

    # region methods

    def get_cell_values(self) -> mCellVals.CellValues:
        """
        Gets cell values

        Returns:
            CellValues: Cell Values instance.
        """
        return mCellVals.CellValues.from_cell(self.get_cell_address())

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
            RangeObj: Range Object
        """
        return mRngObj.RangeObj(
            col_start=self.col, col_end=self.col, row_start=self.row, row_end=self.row, sheet_idx=self.sheet_idx
        )

    # endregion methods

    # region dunder methods

    def __str__(self) -> str:
        return f"{self.col}{self.row}"

    def __eq__(self, other: object) -> bool:
        if isinstance(other, CellObj):
            return self.sheet_idx == other.sheet_idx and self.col == other.col and self.row == other.row
        if isinstance(other, str):
            return str(self) == other.upper()
        return False

    def __add__(self, other: object) -> Self:
        try:
            if isinstance(other, str):
                # string mean add column
                col = cast(mCol.ColObj, self.col_info + other)
                return CellObj(col=col.value, row=self.row, sheet_idx=self.sheet_idx, range_obj=self.range_obj)
            if isinstance(other, mCol.ColObj):
                return CellObj(col=other.value, row=self.row, sheet_idx=self.sheet_idx, range_obj=self.range_obj)
            if isinstance(other, mRow.RowObj):
                return CellObj(col=self.col, row=other.value, sheet_idx=self.sheet_idx, range_obj=self.range_obj)
            if isinstance(other, int):
                return CellObj(col=self.col, row=self.row + other, sheet_idx=self.sheet_idx, range_obj=self.range_obj)
            if isinstance(other, CellObj):
                # add row and column:
                col = cast(mCol.ColObj, self.col_info + other.col_info)
                row = cast(mRow.RowObj, self.row_info + other.row_info)
                return CellObj(col=col.value, row=row.value, sheet_idx=self.sheet_idx, range_obj=self.range_obj)
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            return NotImplemented

    def __sub__(self, other: object) -> Self:
        try:
            if isinstance(other, str):
                # string mean subtract column
                col = cast(mCol.ColObj, self.col_info - other)
                return CellObj(col=col.value, row=self.row, sheet_idx=self.sheet_idx, range_obj=self.range_obj)
            if isinstance(other, mCol.ColObj):
                return CellObj(col=other.value, row=self.row, sheet_idx=self.sheet_idx, range_obj=self.range_obj)
            if isinstance(other, mRow.RowObj):
                return CellObj(col=self.col, row=other.value, sheet_idx=self.sheet_idx, range_obj=self.range_obj)
            if isinstance(other, int):
                return CellObj(col=self.col, row=self.row - other, sheet_idx=self.sheet_idx, range_obj=self.range_obj)
            if isinstance(other, CellObj):
                # subbtract row and column:
                col = cast(mCol.ColObj, self.col_info - other.col_info)
                row = cast(mRow.RowObj, self.row_info - other.row_info)
                return CellObj(col=col.value, row=row.value, sheet_idx=self.sheet_idx, range_obj=self.range_obj)
        except AssertionError as e:
            raise IndexError from e
        except Exception:
            return NotImplemented

    # endregion dunder methods

    # region properties

    @property
    def col_info(self) -> mCol.ColObj:
        """Gets Column Info"""
        try:
            inf = self._col_info
            if inf() is None:
                raise AttributeError
            return inf()
        except AttributeError:
            obj = mCol.ColObj(value=self.col, cell_obj=self)
            object.__setattr__(self, "_col_info", ref(obj))
        return self._col_info()

    @property
    def row_info(self) -> mRow.RowObj:
        """Gets Row Info"""
        try:
            inf = self._row_info
            if inf() is None:
                raise AttributeError
            return inf()
        except AttributeError:
            obj = mRow.RowObj(value=self.row, cell_obj=self)
            object.__setattr__(self, "_row_info", ref(obj))
        return self._row_info()

    @property
    def right(self) -> CellObj:
        """Gets the cell to the right of current cell"""
        try:
            co = self._cell_right
            if co() is None:
                raise AttributeError
            return co()
        except AttributeError:
            co = CellObj(
                col=self.col_info.next.value, row=self.row, sheet_idx=self.sheet_idx, range_obj=self.range_obj
            )
            object.__setattr__(self, "_cell_right", ref(co))
        return self._cell_right()

    @property
    def left(self) -> CellObj:
        """
        Gets the cell to the left of current cell

        Raises:
            IndexError: If cell left is out of range
        """
        try:
            co = self._cell_left
            if co() is None:
                raise AttributeError
            return co()
        except AttributeError:
            try:
                co = CellObj(
                    col=self.col_info.prev.value, row=self.row, sheet_idx=self.sheet_idx, range_obj=self.range_obj
                )
                object.__setattr__(self, "_cell_left", ref(co))
            except IndexError:
                raise
            except AssertionError as e:
                raise IndexError from e
        return self._cell_left()

    @property
    def down(self) -> CellObj:
        """Gets the cell below of current cell"""
        try:
            co = self._cell_down
            if co() is None:
                raise AttributeError
            return co()
        except AttributeError:
            co = CellObj(
                col=self.col, row=self.row_info.next.value, sheet_idx=self.sheet_idx, range_obj=self.range_obj
            )
            object.__setattr__(self, "_cell_down", ref(co))
        return self._cell_down()

    @property
    def up(self) -> CellObj:
        """
        Gets the cell above of current cell

        Raises:
            IndexError: If cell above is out of range
        """
        try:
            co = self._cell_up
            if co() is None:
                raise AttributeError
            return co()
        except AttributeError:
            try:
                co = CellObj(
                    col=self.col, row=self.row_info.prev.value, sheet_idx=self.sheet_idx, range_obj=self.range_obj
                )
                object.__setattr__(self, "_cell_up", ref(co))
            except IndexError:
                raise
            except AssertionError as e:
                raise IndexError from e
        return self._cell_up()

    # endregion properties
