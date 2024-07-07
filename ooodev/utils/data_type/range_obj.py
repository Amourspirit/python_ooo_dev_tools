# region imports
from __future__ import annotations
import contextlib
from typing import Any, TYPE_CHECKING, Generator, cast, overload
from dataclasses import dataclass, field
from weakref import ref
import uno

from ooodev.events.event_singleton import _Events
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.loader import lo as mLo
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.exceptions import ex as mEx
from ooodev.utils import table_helper as mTb
from ooodev.utils.decorator import enforce
from ooodev.loader.inst.doc_type import DocType


if TYPE_CHECKING:
    from com.sun.star.table import CellAddress
    from ooo.dyn.table.cell_range_address import CellRangeAddress
    from ooodev.utils.data_type import cell_values as mCellVals
    from ooodev.calc.calc_doc import CalcDoc

    # from com.sun.star.table import CellRangeAddress
# endregion imports


@enforce.enforce_types
@dataclass(frozen=True)
class RangeObj:
    """
    Range Parts

    .. seealso::
        - :ref:`help_ooodev.utils.data_type.range_obj.RangeObj`

    .. versionchanged:: 0.32.0
        Added support for ``__contains__`` and ``__iter__`` methods. If sheet_idx is set to -2 then no attempt is made to get the sheet index or name from spreadsheet.

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
    """
    Sheet index that this cell value belongs to.
    If value is ``-1`` then the active spreadsheet, if available, is used to get the sheet index.
    If the value is ``-2`` then no sheet index is applied and sheet name will always return and empty string.
    """

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
        if self.sheet_idx == -1:
            with contextlib.suppress(Exception):
                # pylint: disable=no-member
                if mLo.Lo.is_loaded and mLo.Lo.current_doc.DOC_TYPE == DocType.CALC:
                    doc = cast("CalcDoc", mLo.Lo.current_doc)
                    sheet = doc.get_active_sheet()
                    idx = sheet.get_sheet_index()
                    name = sheet.name
                    object.__setattr__(self, "sheet_idx", idx)
                    object.__setattr__(self, "_sheet_name", name)
        if self.sheet_idx < -1:
            object.__setattr__(self, "_sheet_name", "")

    # endregion init

    def __len__(self) -> int:
        """
        Get the number of cells in the range.

        Returns:
            int: Number of cells in range.
        """
        return self.cell_count

    def __copy__(self) -> RangeObj:
        return RangeObj(
            col_start=self.col_start,
            col_end=self.col_end,
            row_start=self.row_start,
            row_end=self.row_end,
            sheet_idx=self.sheet_idx,
        )

    # region methods

    def copy(self) -> RangeObj:
        """
        Copy the current instance.

        Returns:
            RangeObj: New instance of RangeObj

        .. versionadded:: 0.47.5
        """
        return self.__copy__()

    def set_sheet_index(self, idx: int | None = None) -> RangeObj:
        """
        Set the sheet index for the range.

        If ``idx`` is ``None`` then the active sheet index is used.

        Setting the sheet index to -2 will cause the sheet name to always return an empty string.

        Changing the sheet index will cause the sheet name to be re-evaluated.

        Args:
            idx (int, optional): Sheet index, Default ``None``.

        Returns:
            RangeObj: Self

        .. versionadded:: 0.32.0
        """
        if idx is None:
            try:
                # pylint: disable=no-member
                if mLo.Lo.is_loaded and mLo.Lo.current_doc.DOC_TYPE == DocType.CALC:
                    doc = cast("CalcDoc", mLo.Lo.current_doc)
                    sheet = doc.get_active_sheet()
                    idx = sheet.get_sheet_index()
                    name = sheet.name
                    object.__setattr__(self, "sheet_idx", idx)
                    object.__setattr__(self, "_sheet_name", name)
            except Exception:
                object.__setattr__(self, "sheet_idx", -1)
                if hasattr(self, "_sheet_name"):
                    object.__delattr__(self, "_sheet_name")
            return self

        if idx != self.sheet_idx:
            object.__setattr__(self, "sheet_idx", idx)
            if hasattr(self, "_sheet_name"):
                object.__delattr__(self, "_sheet_name")
        return self

    # region from_range()

    @overload
    @staticmethod
    def from_range(range_val: str) -> RangeObj: ...

    @overload
    @staticmethod
    def from_range(range_val: mRngValues.RangeValues) -> RangeObj: ...

    @overload
    @staticmethod
    def from_range(range_val: CellRangeAddress) -> RangeObj: ...

    @staticmethod
    def from_range(range_val: str | mRngValues.RangeValues | CellRangeAddress) -> RangeObj:
        """
        Gets a ``RangeObj`` from are range name.

        Args:
            range_name (str): Range name such as ``A1:D34``.

        Returns:
            RangeObj: Object that represents the name range.
        """

        def handel_event(args: CancelEventArgs, idx: int) -> tuple:

            ret_val = None
            if args.cancel:
                if args.handled is False:
                    args.set("initial_event", "before_style_font_effect")
                    _Events().trigger(GblNamedEvent.EVENT_CANCELED, args)

                if args.handled is False:
                    raise mEx.CancelEventError(args, "Operation canceled")

                if "result" in args.event_data:
                    ret_val = args.event_data["result"]
                else:
                    raise mEx.CancelEventError(args, "Operation canceled, no result data to return.")
            sheet_idx = int(cargs.event_data.get("sheet_index", idx))
            return (sheet_idx, ret_val)

        if hasattr(range_val, "typeName") and getattr(range_val, "typeName") == "com.sun.star.table.CellRangeAddress":
            rng = mRngValues.RangeValues.from_range(cast("CellRangeAddress", range_val))
        else:
            rng = range_val

        # return mTb.TableHelper.get_range_obj(range_name=str(range_val))
        sheet_idx = -2
        cargs = CancelEventArgs("RangeObj.from_range")
        event_data = {"range_val": range_val, "sheet_index": sheet_idx}
        cargs.event_data = event_data
        if isinstance(rng, mRngValues.RangeValues):
            col_start = mTb.TableHelper.make_column_name(rng.col_start, True)
            col_end = mTb.TableHelper.make_column_name(rng.col_end, True)
            row_start = rng.row_start + 1
            row_end = rng.row_end + 1

            cargs.event_data["sheet_index"] = rng.sheet_idx
            _Events().trigger(GblNamedEvent.RANGE_OBJ_BEFORE_FROM_RANGE, cargs)
            sheet_idx, result = handel_event(cargs, rng.sheet_idx)
            if result is not None:
                return result

        else:
            parts = mTb.TableHelper.get_range_parts(str(rng))
            col_start = parts.col_start
            col_end = parts.col_end
            row_start = parts.row_start
            row_end = parts.row_end
            sheet_name = parts.sheet
            if sheet_name:
                sheet_idx = -1
                cargs.event_data["sheet_index"] = sheet_idx

            _Events().trigger(GblNamedEvent.RANGE_OBJ_BEFORE_FROM_RANGE, cargs)
            sheet_idx, result = handel_event(cargs, sheet_idx)
            if result is not None:
                return result

            if sheet_idx == -1:
                with contextlib.suppress(Exception):
                    # pylint: disable=no-member
                    if mLo.Lo.is_loaded and mLo.Lo.current_doc.DOC_TYPE == DocType.CALC:
                        doc = cast("CalcDoc", mLo.Lo.current_doc)
                        sheet = doc.get_sheet(sheet_name=sheet_name)
                        sheet_idx = sheet.get_sheet_index()

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
            s = f"{self.sheet_name}.{s}"
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

    # region get_col()

    @overload
    def get_col(self, col: int) -> RangeObj:
        """
        Gets a Range object that represents only the column range.

        Args:
            col (int): Zero-based column index.

        Returns:
            RangeObj: Range Object.
        """
        ...

    @overload
    def get_col(self, col: str) -> RangeObj:
        """
        Gets a Range object that represents only the column range.

        Args:
            col (str): Column Letter such as ``A``.

        Returns:
            RangeObj: Range Object.
        """
        ...

    def get_col(self, col: int | str) -> RangeObj:
        """
        Gets a Range object that represents only the column range.

        Args:
            col (int | str): Zero-based column index or column letter such as ``A``.

        Raises:
            IndexError: If column index is out of range.

        Returns:
            RangeObj: Range Object


        .. versionadded:: 0.15.1
        """

        if isinstance(col, str):
            col_name = col.upper()
            idx = mTb.TableHelper.col_name_to_int(col_name, True)
        else:
            idx = int(col)
            col_name = mTb.TableHelper.make_column_name(idx, True)
        if idx < self.start_col_index or idx > self.end_col_index:
            raise IndexError(f"Column index {idx} is out of range")

        return RangeObj(
            col_start=col_name,
            col_end=col_name,
            row_start=self.row_start,
            row_end=self.row_end,
            sheet_idx=self.sheet_idx,
        )

    # endregion get_col()

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

    def get_row(self, row: int) -> RangeObj:
        """
        Gets a Range object that represents a row in range.

        Args:
            row (int): Zero-based row index.

        Raises:
            IndexError: If row index is out of range.

        Returns:
            RangeObj: Range Object


        .. versionadded:: 0.15.1
        """
        if row < self.start_row_index or row > self.end_row_index:
            raise IndexError(f"Row index {row} is out of range")

        return RangeObj(
            col_start=self.col_start,
            col_end=self.col_end,
            row_start=row + 1,
            row_end=row + 1,
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
    def contains(self, cell_obj: mCellObj.CellObj) -> bool: ...

    @overload
    def contains(self, cell_addr: CellAddress) -> bool: ...

    @overload
    def contains(self, cell_vals: mCellVals.CellValues) -> bool: ...

    @overload
    def contains(self, cell_name: str) -> bool: ...

    def contains(self, *args, **kwargs) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            cell_obj (CellObj): Cell object
            cell_addr (CellAddress): Cell address
            cell_vals (CellValues): Cell Values
            cell_name (str): Cell name

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.

        Note:
            If cell input contains sheet info the it is use in comparison.
            Otherwise sheet is ignored.

        See Also:
            - :ref:`help_ooodev.utils.data_type.range_obj.RangeObj.contains`
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
            plus = self.start_col_index + self.col_count + 1
            for i in range(self.start_col_index + 1, plus):
                col_name = mTb.TableHelper.make_column_name(i)
                yield mCellObj.CellObj(
                    col=col_name,
                    row=start_cell.row,
                    sheet_idx=start_cell.sheet_idx,
                    range_obj=self,
                )

        def row_gen():
            curr_row = self.row_start
            while curr_row <= self.row_end:
                start_cell = mCellObj.CellObj(
                    col=self.col_start, row=curr_row, sheet_idx=self.sheet_idx, range_obj=self
                )
                yield row_cell_gen(start_cell)
                curr_row += 1

        return row_gen()

    def __getitem__(self, key: str) -> Any:
        """
        Get a cell object from range

        Args:
            key (int): Zero-based index of cell.

        Raises:
            TypeError: If index is not a string.
            IndexError: If index is out of range.

        Returns:
            CellObj: Cell Object

        Example:
            .. code-block:: python

                >>> rng = RangeObj.from_range("A1:C4")
                >>> cell = rng["B3"]
                >>> print(cell)
                B3
        """
        if not isinstance(key, str):
            raise TypeError("Index must be a string that represents a cell name such as 'A1' or 'B2'")

        if not self.contains(key):
            raise IndexError(f"Index '{key}' is out of range")
        parts = mTb.TableHelper.get_cell_parts(key)
        return mCellObj.CellObj(
            col=parts.col,
            row=parts.row,
            range_obj=self,
        )

    def __contains__(self, value: Any) -> bool:
        """
        Gets if current instance contains a cell value.

        Args:
            value (CellObj): Cell object
            value (CellAddress): Cell address
            value (CellValues): Cell Values
            value (str): Cell name

        Returns:
            bool: ``True`` if instance contains cell; Otherwise, ``False``.

        Note:
            If cell input contains sheet info the it is use in comparison.
            Otherwise sheet is ignored.

        See Also:
            - :ref:`help_ooodev.utils.data_type.range_obj.RangeObj.contains`

        .. versionadded:: 0.32.0
        """
        return self.contains(value)

    def __iter__(self) -> Generator[mCellObj.CellObj, None, None]:
        """
        Iterates over all cells in the range.

        The iteration is done in a column-major order, meaning that the cells are iterated over by column, then by row.

        Example:
            .. code-block:: python

                # each cell is an instance of CellObj
                >>> rng = RangeObj.from_range("A1:C4")
                >>> for cell in rng:
                >>>     print(cell)
                A1
                B1
                C1
                A2
                B2
                C2
                A3
                B3
                C3
                A4
                B4
                C4

        Yields:
            Generator[mCellObj.CellObj, None, None]: Cell Object

        See Also:
            - :ref:`help_ooodev.utils.data_type.range_obj.RangeObj.__iter__`

        .. versionadded:: 0.32.0
        """
        # return iter(self.get_cells())
        for cells in self.get_cells():
            yield from cells

    def __str__(self) -> str:
        """
        Convert range to string

        Returns:
            str: Inf the format of ``A1:C4``
        """
        return f"{self.col_start}{self.row_start}:{self.col_end}{self.row_end}"

    def __eq__(self, other: object) -> bool:
        """
        Compare if two ranges are equal

        Args:
            other (object): Range Object, ``RangeValues`` or str in range format such as ``A1:C4``

        Returns:
            bool: ``True`` if equal; Otherwise, ``False``
        """
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
        """
        Add range to another range.

        Args:
            other (object): Other Range, Row, Column, Cell or str in range format such as ``A1:C4``

        Returns:
            RangeObj: _description_
        """
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

    def __truediv__(self, other: object) -> RangeObj:
        rng_obj: RangeObj | None = None
        if isinstance(other, RangeObj):
            rng_obj = other
        if isinstance(other, str):
            try:
                parts = mTb.TableHelper.get_range_parts(other)
                rng_obj = RangeObj(
                    col_start=parts.col_start,
                    col_end=parts.col_end,
                    row_start=parts.row_start,
                    row_end=parts.row_end,
                    sheet_idx=0,
                )
            except Exception as e:
                raise ValueError(f'String Value "{other}" cannot be converted to a RangeObj') from e

        if rng_obj is not None:
            row_start = min(self.row_start, rng_obj.row_start)
            row_end = max(self.row_end, rng_obj.row_end)
            if self.start_col_index < rng_obj.start_col_index:
                col_start = self.col_start
            else:
                col_start = rng_obj.col_start
            if self.end_col_index > rng_obj.end_col_index:
                col_end = self.col_end
            else:
                col_end = rng_obj.col_end
            return RangeObj(
                col_start=col_start, col_end=col_end, row_start=row_start, row_end=row_end, sheet_idx=self.sheet_idx
            )
        return NotImplemented

    def __rtruediv__(self, other):
        return self.__truediv__(other)

    # endregion methods

    # region properties

    @property
    def sheet_name(self) -> str:
        """Gets sheet name"""
        # return self._sheet_name
        try:
            return self._sheet_name  # type: ignore
        except AttributeError:
            name = ""
            if self.sheet_idx < 0:
                return name
            with contextlib.suppress(Exception):
                # pylint: disable=no-member
                if mLo.Lo.is_loaded and mLo.Lo.current_doc.DOC_TYPE == DocType.CALC:
                    doc = cast("CalcDoc", mLo.Lo.current_doc)
                    sheet = doc.sheets[self.sheet_idx]
                    name = sheet.name
                    object.__setattr__(self, "_sheet_name", name)
        return name

    @property
    def cell_start(self) -> mCellObj.CellObj:
        """Gets the Start Cell object for Range"""
        # pylint: disable=no-member
        try:
            co = self._cell_start  # type: ignore
            if co() is None:
                raise AttributeError
            return co()
        except AttributeError:
            c = mCellObj.CellObj(col=self.col_start, row=self.row_start, sheet_idx=self.sheet_idx, range_obj=self)
            object.__setattr__(self, "_cell_start", ref(c))
        return self._cell_start()  # type: ignore

    @property
    def cell_end(self) -> mCellObj.CellObj:
        """Gets the End Cell object for Range"""
        # pylint: disable=no-member
        try:
            co = self._cell_end  # type: ignore
            if co() is None:
                raise AttributeError
            return co()
        except AttributeError:
            c = mCellObj.CellObj(col=self.col_end, row=self.row_end, sheet_idx=self.sheet_idx, range_obj=self)
            object.__setattr__(self, "_cell_end", ref(c))
        return self._cell_end()  # type: ignore

    @property
    def start_row_index(self) -> int:
        """Gets start row zero-based index"""
        return self.row_start - 1

    @property
    def start_col_index(self) -> int:
        """Gets start column zero-based index"""
        # pylint: disable=no-member
        try:
            return self._start_col_index  # type: ignore
        except AttributeError:
            object.__setattr__(self, "_start_col_index", self.cell_start.col_obj.index)
        return self._start_col_index  # type: ignore

    @property
    def end_row_index(self) -> int:
        """Gets end row zero-based index"""
        return self.row_end - 1

    @property
    def end_col_index(self) -> int:
        """Gets end column zero-based index"""
        # pylint: disable=no-member
        try:
            return self._end_col_index  # type: ignore
        except AttributeError:
            object.__setattr__(self, "_end_col_index", self.cell_end.col_obj.index)
        return self._end_col_index  # type: ignore

    @property
    def row_count(self) -> int:
        """
        Gets the number of rows in the current range

        Returns:
            int: Number of rows
        """
        start = self.start_row_index
        end = self.end_row_index
        return abs(end - start) + 1

    @property
    def col_count(self) -> int:
        """
        Gets the number of columns in the current range

        Returns:
            int: Number of columns
        """
        start = self.start_col_index
        end = self.end_col_index
        return abs(end - start) + 1

    @property
    def cell_count(self) -> int:
        """
        Gets the number of cell in the current range

        Returns:
            int: Number of cells
        """
        return self.row_count * self.col_count

    # endregion properties


from ooodev.utils.data_type import row_obj as mRowObj
from ooodev.utils.data_type import col_obj as mColObj
from ooodev.utils.data_type import cell_obj as mCellObj
from ooodev.utils.data_type import range_values as mRngValues
