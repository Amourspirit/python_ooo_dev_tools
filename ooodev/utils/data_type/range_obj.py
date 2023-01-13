# region imports
from __future__ import annotations
from typing import TYPE_CHECKING, Generator, overload
from dataclasses import dataclass, field
from weakref import ref
from .. import lo as mLo
from .. import table_helper as mTb
from ...office import calc as mCalc
from ..decorator import enforce


import uno
from ooo.dyn.table.cell_range_address import CellRangeAddress

if TYPE_CHECKING:
    from . import cell_values as mCellVals
    from com.sun.star.table import CellAddress
# endregion imports


@enforce.enforce_types
@dataclass(frozen=True)
class RangeObj:
    """
    Range Parts

    .. versionadded:: 0.8.2
    """

    # region init

    col_start: str
    """Column start such as ``A``"""
    col_end: str
    """Column end such as ``C``"""
    row_start: int
    """Row start such as ``1``"""
    row_end: int
    """Row end such as ``125``"""
    start: mCellObj.CellObj = field(init=False, repr=False, hash=False)
    """Start Cell Object"""
    end: mCellObj.CellObj = field(init=False, repr=False, hash=False)
    """End Cell Object"""
    sheet_idx: int = -1
    """Sheet index that this range value belongs to"""

    def __post_init__(self):
        row_start = self.row_start
        row_end = self.row_end
        if row_start > row_end:
            row_start, row_end = row_end, row_start
            object.__setattr__(self, "row_start", row_start)
            object.__setattr__(self, "row_end", row_end)

        object.__setattr__(self, "col_start", self.col_start.upper())
        object.__setattr__(self, "col_end", self.col_end.upper())

        col_start_num = mTb.TableHelper.col_name_to_int(self.col_start)
        col_end_num = mTb.TableHelper.col_name_to_int(self.col_end)
        if col_start_num > col_end_num:
            # swap columns
            col_start, col_end = self.col_end, self.col_start
            object.__setattr__(self, "col_start", col_start)
            object.__setattr__(self, "col_end", col_end)

        start = mCellObj.CellObj(col=self.col_start, row=self.row_start, range_obj=self)
        end = mCellObj.CellObj(col=self.col_end, row=self.row_end, range_obj=self)

        object.__setattr__(self, "start", start)
        object.__setattr__(self, "end", end)
        if self.sheet_idx < 0:
            try:
                if mLo.Lo.is_loaded:
                    idx = mCalc.Calc.get_sheet_index()
                    object.__setattr__(self, "sheet_idx", idx)
            except:
                pass

    # endregion init

    # region methods

    # region from_range()

    @overload
    @staticmethod
    def from_range(range_val: str) -> RangeObj:
        ...

    @overload
    @staticmethod
    def from_range(range_val: mRngValues.RangeValues) -> RangeObj:
        ...

    @overload
    @staticmethod
    def from_range(range_val: CellRangeAddress) -> RangeObj:
        ...

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
            if sheet_name and mLo.Lo.is_loaded:
                sheet = mCalc.Calc.get_sheet(doc=mCalc.Calc.get_current_doc(), sheet_name=sheet_name)
                sheet_idx = mCalc.Calc.get_sheet_index(sheet)

        return RangeObj(
            col_start=col_start, col_end=col_end, row_start=row_start, row_end=row_end, sheet_idx=sheet_idx
        )

    # endregion from_range()

    def get_range_values(self) -> mRngValues.RangeValues:
        """
        Gets ``RangeValues``

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

    def get_cells(self) -> Generator[Generator[mCellObj.CellObj, None, None], None, None]:
        """
        Get a generator that can loop over all cells within current range.

        Returns:
            Generator[Generator[CellObj, None, None], None, None]: Nested Generator that returns :py:class:`~.cell_obj.CellObj` for each cell in current range.

        Example:
            .. code-block:: python

                >>> rng = RangeObj.from_range("A1:C3")
                >>> for row, cells in enumerate(rng.get_cells()):
                >>>     print("Row:", row)
                >>>     for cell in cells:
                >>>         print(f"  {cell}: {repr(cell)}")
                >>> print("Done")
                Row: 0
                  A1: CellObj(col='A', row=1, sheet_idx=0)
                  B1: CellObj(col='B', row=1, sheet_idx=0)
                  C1: CellObj(col='C', row=1, sheet_idx=0)
                Row: 1
                  A2: CellObj(col='A', row=2, sheet_idx=0)
                  B2: CellObj(col='B', row=2, sheet_idx=0)
                  C2: CellObj(col='C', row=2, sheet_idx=0)
                Row: 2
                  A3: CellObj(col='A', row=3, sheet_idx=0)
                  B3: CellObj(col='B', row=3, sheet_idx=0)
                  C3: CellObj(col='C', row=3, sheet_idx=0)
                Done
        """

        def row_cell_gen(start_cell: mCellObj.CellObj):
            # idx 17, 10
            cplus = self.start_col_index + self.col_count + 1
            for i in range(self.start_col_index + 1, cplus):
                col_name = mTb.TableHelper.make_column_name(i)
                cell = mCellObj.CellObj(
                    col=col_name, row=start_cell.row, sheet_idx=start_cell.sheet_idx, range_obj=self
                )
                yield cell

        def row_gen():
            curr_row = self.row_start
            while curr_row <= self.row_end:
                start_cell = mCellObj.CellObj(
                    col=self.col_start, row=curr_row, sheet_idx=self.sheet_idx, range_obj=self
                )
                yield row_cell_gen(start_cell)
                curr_row += 1

        return row_gen()

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

    def __add__(self, other: object) -> RangeObj:
        if isinstance(other, str):
            # add cols to right of range
            cols = mTb.TableHelper.col_name_to_int(other)
            current_rv = self.get_range_values()
            rv = current_rv.add_cols(cols)
            return RangeObj.from_range(rv)
        if isinstance(other, int):
            # add rows to bottom of range
            current_rv = self.get_range_values()
            rv = current_rv.add_rows(other)
            return RangeObj.from_range(rv)
        if isinstance(other, mRowObj.RowObj):
            current_rv = self.get_range_values()
            rv = current_rv.add_rows(other.value)
            return RangeObj.from_range(rv)
        if isinstance(other, mColObj.ColObj):
            current_rv = self.get_range_values()
            rv = current_rv.add_cols(other.index + 1)
            return RangeObj.from_range(rv)
        if isinstance(other, mCellObj.CellObj):
            current_rv = self.get_range_values()
            rv = current_rv.add_cols(other.col_obj.index + 1)
            rv = rv.add_rows(other.row)
            return RangeObj.from_range(rv)
        return NotImplemented

    def __radd__(self, other: object) -> RangeObj:
        if isinstance(other, str):
            # add cols to left of range
            cols = mTb.TableHelper.col_name_to_int(other)
            current_rv = self.get_range_values()
            rv = current_rv.add_cols(cols, False)
            return RangeObj.from_range(rv)
        if isinstance(other, int):
            # add rows to top of range
            current_rv = self.get_range_values()
            rv = current_rv.add_rows(other, False)
            return RangeObj.from_range(rv)
        if isinstance(other, mRowObj.RowObj):
            current_rv = self.get_range_values()
            rv = current_rv.add_rows(other.value, False)
            return RangeObj.from_range(rv)
        if isinstance(other, mColObj.ColObj):
            current_rv = self.get_range_values()
            rv = current_rv.add_cols(other.index + 1, False)
            return RangeObj.from_range(rv)
        if isinstance(other, mCellObj.CellObj):
            current_rv = self.get_range_values()
            rv = current_rv.add_cols(other.col_obj.index + 1, False)
            rv = rv.add_rows(other.row, False)
            return RangeObj.from_range(rv)
        return NotImplemented

    def __sub__(self, other: object) -> RangeObj:
        if isinstance(other, str):
            # subtract col from right of range
            cols = mTb.TableHelper.col_name_to_int(other)
            current_rv = self.get_range_values()
            rv = current_rv.subtract_cols(cols)
            return RangeObj.from_range(rv)
        if isinstance(other, int):
            # subtract rows from bottom of range
            current_rv = self.get_range_values()
            rv = current_rv.subtract_rows(other)
            return RangeObj.from_range(rv)
        if isinstance(other, mRowObj.RowObj):
            current_rv = self.get_range_values()
            rv = current_rv.subtract_rows(other.value)
            return RangeObj.from_range(rv)
        if isinstance(other, mColObj.ColObj):
            current_rv = self.get_range_values()
            rv = current_rv.subtract_cols(other.index + 1)
            return RangeObj.from_range(rv)
        if isinstance(other, mCellObj.CellObj):
            current_rv = self.get_range_values()
            rv = current_rv.subtract_cols(other.col_obj.index + 1)
            rv = rv.subtract_rows(other.row)
            return RangeObj.from_range(rv)
        return NotImplemented

    def __rsub__(self, other: object) -> RangeObj:
        if isinstance(other, str):
            # subtract col from left of range
            cols = mTb.TableHelper.col_name_to_int(other)
            current_rv = self.get_range_values()
            rv = current_rv.subtract_cols(cols, False)
            return RangeObj.from_range(rv)
        if isinstance(other, int):
            # subtract rows from top of range
            current_rv = self.get_range_values()
            rv = current_rv.subtract_rows(other, False)
            return RangeObj.from_range(rv)
        if isinstance(other, mRowObj.RowObj):
            current_rv = self.get_range_values()
            rv = current_rv.subtract_rows(other.value, False)
            return RangeObj.from_range(rv)
        if isinstance(other, mColObj.ColObj):
            current_rv = self.get_range_values()
            rv = current_rv.subtract_cols(other.index + 1, False)
            return RangeObj.from_range(rv)
        if isinstance(other, mCellObj.CellObj):
            current_rv = self.get_range_values()
            rv = current_rv.subtract_cols(other.col_obj.index + 1, False)
            rv = rv.subtract_rows(other.row, False)
            return RangeObj.from_range(rv)
        return NotImplemented

    # endregion methods

    # region properties

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
    def cell_start(self) -> mCellObj.CellObj:
        """Gets the Start Cell object for Range"""
        try:
            co = self._cell_start
            if co() is None:
                raise AttributeError
            return co()
        except AttributeError:
            c = mCellObj.CellObj(col=self.col_start, row=self.row_start, sheet_idx=self.sheet_idx, range_obj=self)
            object.__setattr__(self, "_cell_start", ref(c))
        return self._cell_start()

    @property
    def cell_end(self) -> mCellObj.CellObj:
        """Gets the End Cell object for Range"""
        try:
            co = self._cell_end
            if co() is None:
                raise AttributeError
            return co()
        except AttributeError:
            c = mCellObj.CellObj(col=self.col_end, row=self.row_end, sheet_idx=self.sheet_idx, range_obj=self)
            object.__setattr__(self, "_cell_end", ref(c))
        return self._cell_end()

    @property
    def start_row_index(self) -> int:
        """Gets start row zero-based index"""
        return self.row_start - 1

    @property
    def start_col_index(self) -> int:
        """Gets start column zero-based index"""
        try:
            return self._start_col_index
        except AttributeError:
            object.__setattr__(self, "_start_col_index", self.cell_start.col_obj.index)
        return self._start_col_index

    @property
    def end_row_index(self) -> int:
        """Gets end row zero-based index"""
        return self.row_end - 1

    @property
    def end_col_index(self) -> int:
        """Gets end column zero-based index"""
        try:
            return self._end_col_index
        except AttributeError:
            object.__setattr__(self, "_end_col_index", self.cell_end.col_obj.index)
        return self._end_col_index

    @property
    def row_count(self) -> int:
        """
        Gets the number of rows in the current range

        Returns:
            int: Number of rows
        """
        start = self.start_row_index
        end = self.end_row_index
        count = abs(end - start) + 1
        return count

    @property
    def col_count(self) -> int:
        """
        Gets the number of columns in the current range

        Returns:
            int: Number of columns
        """
        start = self.start_col_index
        end = self.end_col_index
        count = abs(end - start) + 1
        return count

    @property
    def cell_count(self) -> int:
        """
        Gets the number of cell in the current range

        Returns:
            int: Number of cells
        """
        return self.row_count * self.col_count

    # endregion properties


from . import row_obj as mRowObj
from . import col_obj as mColObj
from . import cell_obj as mCellObj
from . import range_values as mRngValues
