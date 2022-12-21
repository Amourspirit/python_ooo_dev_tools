from __future__ import annotations
from typing import TYPE_CHECKING, overload
from dataclasses import dataclass, field
from . import cell_obj as mCo
from . import range_values as mRngValues
from . import cell_obj as mCell
from .. import table_helper as mTb
from ...office import calc as mCalc
from ..decorator import enforce

import uno
from ooo.dyn.table.cell_range_address import CellRangeAddress

if TYPE_CHECKING:
    from . import cell_obj as mCellObj
    from . import cell_values as mCellVals
    from com.sun.star.table import CellAddress


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
    sheet_idx: int = -1
    """Sheet index that this range value belongs to"""

    def __post_init__(self):
        object.__setattr__(self, "col_start", self.col_start.upper())
        object.__setattr__(self, "col_end", self.col_end.upper())

        object.__setattr__(self, "start", mCo.CellObj.from_cell(f"{self.col_start}{self.row_start}"))
        object.__setattr__(self, "end", mCo.CellObj.from_cell(f"{self.col_end}{self.row_end}"))
        if self.sheet_idx < 0:
            try:
                idx = mCalc.Calc.get_sheet_index()
                object.__setattr__(self, "sheet_idx", idx)
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
            sheet_idx = rng.sheet_idx
        else:
            parts = mTb.TableHelper.get_range_parts(rng)
            col_start = parts.col_start
            col_end = parts.col_end
            row_start = parts.row_start
            row_end = parts.row_end
            sheet_name = parts.sheet
            sheet_idx = -1
            if sheet_name:
                sheet = mCalc.Calc.get_sheet(doc=mCalc.Calc.get_current_doc(), sheet_name=sheet_name)
                sheet_idx = mCalc.Calc.get_sheet_index(sheet)

        return RangeObj(
            col_start=col_start, row_start=row_start, col_end=col_end, row_end=row_end, sheet_idx=sheet_idx
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

    def get_start_col(self) -> RangeObj:
        """
        Gets a Range object that represents only the start column range.

        Returns:
            RangeObj: Range Object
        """
        return RangeObj(
            col_start=self.col_start,
            col_end=self.col_start,
            row_start=self.row_start,
            row_end=self.row_end,
            sheet_idx=self.sheet_idx,
        )

    def get_end_col(self) -> RangeObj:
        """
        Gets a Range object that represents only the end column range.

        Returns:
            RangeObj: Range Object
        """
        return RangeObj(
            col_start=self.col_end,
            col_end=self.col_end,
            row_start=self.row_start,
            row_end=self.row_end,
            sheet_idx=self.sheet_idx,
        )

    def get_start_row(self) -> RangeObj:
        """
        Gets a Range object that represents only the start row range.

        Returns:
            RangeObj: Range Object
        """
        return RangeObj(
            col_start=self.col_start,
            col_end=self.col_end,
            row_start=self.row_start,
            row_end=self.row_start,
            sheet_idx=self.sheet_idx,
        )

    def get_end_row(self) -> RangeObj:
        """
        Gets a Range object that represents only the end row range.

        Returns:
            RangeObj: Range Object
        """
        return RangeObj(
            col_start=self.col_start,
            col_end=self.col_end,
            row_start=self.row_end,
            row_end=self.row_end,
            sheet_idx=self.sheet_idx,
        )

    def is_single_col(self) -> bool:
        """
        Gets if instance is a single column or multi-column

        Returns:
            bool: ``True`` if single column; Otherwise, ``False``

        Note:
            If instance is a single cell address then ``True`` is returned.
        """
        cv = mRngValues.RangeValues.from_range(self)
        return cv.is_single_col()

    def is_single_row(self) -> bool:
        """
        Gets if instance is a single row or multi-row

        Returns:
            bool: ``True`` if single row; Otherwise, ``False``

        Note:
            If instance is a single cell address then ``True`` is returned.
        """
        cv = mRngValues.RangeValues.from_range(self)
        return cv.is_single_row()

    def is_single_cell(self) -> bool:
        """
        Gets if a instance is a single cell or a range

        Returns:
            bool: ``True`` if single cell; Otherwise, ``False``
        """
        cv = mRngValues.RangeValues.from_range(self)
        return cv.is_single_cell()

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
        rv = self.get_range_values()
        return rv.contains(*args, **kwargs)

    # endregion contains()

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

    @property
    def sheet_name(self) -> str:
        """Gets sheet name"""
        # return self._sheet_name
        try:
            return self._sheet_name
        except AttributeError:
            name = ""
            if self.sheet_idx < 0:
                return name
            try:
                sheet = mCalc.Calc.get_sheet(doc=mCalc.Calc.get_current_doc(), index=self.sheet_idx)
                name = mCalc.Calc.get_sheet_name(sheet=sheet)
                object.__setattr__(self, "_sheet_name", name)
            except:
                pass
        return name

    @property
    def cell_start(self) -> mCell.CellObj:
        """Gets the Start Cell object for Range"""
        try:
            return self._cell_start
        except AttributeError:
            c = mCell.CellObj(col=self.col_start, row=self.row_start, sheet_idx=self.sheet_idx, range_obj=self)
            object.__setattr__(self, "_cell_start", c)
        return self._cell_start

    @property
    def cell_end(self) -> mCell.CellObj:
        """Gets the End Cell object for Range"""
        try:
            return self._cell_end
        except AttributeError:
            c = mCell.CellObj(col=self.col_end, row=self.row_end, sheet_idx=self.sheet_idx, range_obj=self)
            object.__setattr__(self, "_cell_end", c)
        return self._cell_end
