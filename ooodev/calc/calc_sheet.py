from __future__ import annotations
from typing import Any, List, Tuple, cast, overload, Sequence, TYPE_CHECKING
import uno

from com.sun.star.sheet import XSpreadsheet
from com.sun.star.util import XProtectable
from com.sun.star.sheet import XSheetCellRange

from ooo.dyn.sheet.cell_flags import CellFlagsEnum as CellFlagsEnum


if TYPE_CHECKING:
    from com.sun.star.sheet import XSheetCellCursor
    from com.sun.star.sheet import XScenario
    from com.sun.star.table import CellAddress
    from com.sun.star.table import XCell
    from com.sun.star.table import XCellRange

    from ooo.dyn.beans.property_value import PropertyValue
    from ooo.dyn.table.cell_range_address import CellRangeAddress
    from com.sun.star.util import XSearchable
    from com.sun.star.util import XSearchDescriptor
    from ooodev.proto.style_obj import StyleT
    from ooodev.units import UnitT
    from ooodev.utils.type_var import Row, Column, Table, TupleArray, FloatTable
    from .calc_doc import CalcDoc

from ooodev.utils import lo as mLo
from ooodev.utils import props as mProps
from ooodev.utils.data_type import cell_obj as mCellObj
from ooodev.utils.data_type import range_obj as mRngObj
from ooodev.office import calc as mCalc
from ooodev.adapter.sheet.spreadsheet_comp import SpreadsheetComp
from . import calc_cell_range as mCalcCellRange
from . import calc_cell as mCalcCell
from . import calc_cell_cursor as mCalcCellCursor


class CalcSheet(SpreadsheetComp):
    def __init__(self, owner: CalcDoc, sheet: XSpreadsheet) -> None:
        super().__init__(sheet)  # type: ignore
        self.__owner = owner

    # region get_array()
    @overload
    def get_array(self, *, cell_range: XCellRange) -> TupleArray:
        ...

    @overload
    def get_array(self, *, range_name: str) -> TupleArray:
        ...

    @overload
    def get_array(self, *, range_obj: mRngObj.RangeObj) -> TupleArray:
        ...

    @overload
    def get_array(self, *, cell_obj: mCellObj.CellObj) -> TupleArray:
        ...

    def get_array(self, **kwargs) -> TupleArray:
        """
        Gets Array of data from a spreadsheet.

        Args:
            cell_range (XCellRange): Cell range to get data from.
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range of data to get such as "A1:E16"
            range_obj (RangeObj): Range object
            cell_obj (CellObj): Cell Object

        Raises:
            MissingInterfaceError: if interface is missing

        Returns:
            TupleArray: Resulting data array.
        """
        sheet_names = {"range_name", "range_obj", "cell_obj"}
        if kwargs.keys() & sheet_names:
            kwargs["sheet"] = self.component
        return mCalc.Calc.get_array(**kwargs)

    # endregion get_array()

    # region get_float_array()
    @overload
    def get_float_array(self, *, cell_range: XCellRange) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        Args:
            cell_range (XCellRange): Cell range to get data from.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        ...

    @overload
    def get_float_array(self, *, range_name: str) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        Args:
            range_name (str): Range to get array of floats from such as 'A1:E18'.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        ...

    @overload
    def get_float_array(self, *, range_obj: mRngObj.RangeObj) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        Args:
            range_obj (RangeObj): Range object.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        ...

    @overload
    def get_float_array(self, *, cell_obj: mCellObj.CellObj) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        ...

    def get_float_array(self, **kwargs) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        Args:
            cell_range (XCellRange): Cell range to get data from.
            range_name (str): Range to get array of floats from such as 'A1:E18'.
            range_obj (RangeObj): Range object.
            cell_obj (CellObj): Cell Object.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        sheet_names = {"range_name", "range_obj", "cell_obj"}
        if kwargs.keys() & sheet_names:
            kwargs["sheet"] = self.component
        return mCalc.Calc.get_float_array(**kwargs)

    # endregion get_float_array()

    # region get_cell()
    @overload
    def get_cell(self, *, cell: XCell) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            cell (XCell): Cell

        Returns:
            CalcCell: cell
        """
        ...

    @overload
    def get_cell(self, *, addr: CellAddress) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            addr (CellAddress): Cell Address

        Returns:
            CalcCell: cell
        """
        ...

    @overload
    def get_cell(self, *, cell_name: str) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            cell_name (str): Cell Name such as 'A1'

        Returns:
            CalcCell: cell
        """
        ...

    @overload
    def get_cell(self, *, cell_obj: mCellObj.CellObj) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            cell_obj: (CellObj): Cell object

        Returns:
            CalcCell: cell
        """
        ...

    @overload
    def get_cell(self, *, col: int, row: int) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            col (int): Cell column
            row (int): cell row

        Returns:
            CalcCell: cell
        """
        ...

    @overload
    def get_cell(self, *, cell_range: XCellRange) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            cell_range (XCellRange): Cell Range

        Returns:
            CalcCell: cell
        """
        ...

    @overload
    def get_cell(self, *, cell_range: XCellRange, col: int, row: int) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            cell_range (XCellRange): Cell Range
            col (int): Cell column
            row (int): cell row

        Returns:
            CalcCell: cell
        """
        ...

    def get_cell(self, **kwargs) -> mCalcCell.CalcCell:
        """
        Gets a cell

        Args:
            addr (CellAddress): Cell Address
            cell_name (str): Cell Name such as 'A1'
            cell_obj: (CellObj): Cell object
            cell_range (XCellRange): Cell Range
            col (int): Cell column
            row (int): cell row
            cell (XCell): Cell

        Returns:
            CalcCell: cell
        """
        sheet_names = {"addr", "cell_name", "cell_obj", "col"}
        if kwargs.keys() & sheet_names:
            if not "cell_range" in kwargs:
                kwargs["sheet"] = self.component
        x_cell = mCalc.Calc.get_cell(**kwargs)
        cell_obj = mCalc.Calc.get_cell_obj(cell=x_cell)
        return mCalcCell.CalcCell(owner=self, cell=cell_obj)

    # endregion get_cell()

    # region get_range()
    @overload
    def get_range(self, *, range_name: str) -> mCalcCellRange.CalcCellRange:
        """
        Gets a range Object representing a range.

        Args:
            range_name (str): Cell range as string.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def get_range(self, *, cr_addr: CellRangeAddress) -> mCalcCellRange.CalcCellRange:
        """
        Gets a range Object representing a range.

        Args:
            cr_addr (CellRangeAddress): Cell Range Address.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def get_range(self, *, range_obj: mRngObj.RangeObj) -> mCalcCellRange.CalcCellRange:
        """
        Gets a range Object representing a range.

        Args:
            range_obj (RangeObj): Range Object. If passed in the same RangeObj is returned.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def get_range(self, *, cell_obj: mCellObj.CellObj) -> mCalcCellRange.CalcCellRange:
        """
        Gets a range Object representing a range.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def get_range(self, *, cell_range: XCellRange) -> mCalcCellRange.CalcCellRange:
        """
        Gets a range Object representing a range.

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def get_range(self, *, col_start: int, row_start: int, col_end: int, row_end: int) -> mCalcCellRange.CalcCellRange:
        """
        Gets a range Object representing a range.

        Args:
            col_start (int): Zero-based start column index.
            row_start (int): Zero-based start row index.
            col_end (int): Zero-based end column index.
            row_end (int): Zero-based end row index.

        Returns:
            RangeObj: Range object.
        """
        ...

    def get_range(self, **kwargs) -> mCalcCellRange.CalcCellRange:
        """
        Gets a range Object representing a range.

        Args:
            range_name (str): Cell range as string.
            cell_range (XCellRange): Cell Range.
            cr_addr (CellRangeAddress): Cell Range Address.
            cell_obj (CellObj): Cell Object.
            range_obj (RangeObj): Range Object. If passed in the same RangeObj is returned.
            col_start (int): Zero-based start column index.
            row_start (int): Zero-based start row index.
            col_end (int): Zero-based end column index.
            row_end (int): Zero-based end row index.

        Returns:
            RangeObj: Range object.
        """
        sheet_names = {"cell_range", "col_start"}
        if kwargs.keys() & sheet_names:
            if not "sheet" in kwargs:
                kwargs["sheet"] = self.component
        range_obj = mCalc.Calc.get_range_obj(**kwargs)
        return mCalcCellRange.CalcCellRange(self, range_obj)

    # endregion get_range()

    def get_col_range(self, idx: int) -> mCalcCellRange.CalcCellRange:
        """
        Get Column by index

        Args:
            idx (int): Zero-based column index

        Raises:
            MissingInterfaceError: if unable to find interface

        Returns:
            CalcCellRange: Cell range
        """
        result = mCalc.Calc.get_col_range(self.component, idx)
        return mCalcCellRange.CalcCellRange(self, result)

    # region get_row()
    @overload
    def get_row(self, calc_cell_range: mCalcCellRange.CalcCellRange) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            calc_cell_range (CalcCellRange): Calc cell range to get column data from.

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        ...

    @overload
    def get_row(self, cell_range: XCellRange) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            cell_range (XCellRange): Cell range to get column data from.

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        ...

    @overload
    def get_row(self, row_idx: int) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            row_idx (int): Zero base row index such as `0` for row `1`

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        ...

    @overload
    def get_row(self, range_name: str) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            range_name (str): Range such as 'A1:A12'

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        ...

    @overload
    def get_row(self, cell_obj: mCellObj.CellObj) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            cell_obj (CellObj): Cell Object

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        ...

    @overload
    def get_row(self, range_obj: mRngObj.RangeObj) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            range_obj (RangeObj): Range Object

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        ...

    def get_row(self, *args, **kwargs) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            calc_cell_range (CalcCellRange): Calc cell range to get column data from.
            cell_range (XCellRange): Cell range to get column data from.
            row_idx (int): Zero base row index such as `0` for row `1`
            range_name (str): Range such as 'A1:A12'
            cell_obj (CellObj): Cell Object
            range_obj (RangeObj): Range Object

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        if kwargs:
            if "calc_cell_range" in kwargs:
                arg0 = cast(mCalcCellRange.CalcCellRange, kwargs["calc_cell_range"]).component
                return mCalc.Calc.get_row(arg0)
            elif "cell_range" in kwargs:
                arg0 = cast("XCellRange", kwargs["cell_range"])
                return mCalc.Calc.get_row(arg0)
        return mCalc.Calc.get_row(self.component, *args, **kwargs)

    # endregion get_row()

    def get_row_range(self, idx: int) -> mCalcCellRange.CalcCellRange:
        """
        Get Row by index

        Args:
            sheet (XSpreadsheet): Spreadsheet
            idx (int): Zero-based column index

        Raises:
            MissingInterfaceError: if unable to find interface

        Returns:
            CalcCellRange: Cell range
        """
        result = mCalc.Calc.get_row_range(self.component, idx)
        return mCalcCellRange.CalcCellRange(self, result)

    def get_row_used_first_index(self) -> int:
        """
        Gets the index of the row of the top edge of the used sheet range.

        Returns:
            int: Zero based index of first row used on the sheet.
        """
        return mCalc.Calc.get_row_used_first_index(self.component)

    def get_row_used_last_index(self) -> int:
        """
        Gets the index of the row of the bottom edge of the used sheet range.

        Returns:
            int: Zero based index of last row used on the sheet.
        """
        return mCalc.Calc.get_row_used_last_index(self.component)

    def get_selected_cell(self) -> mCalcCell.CalcCell:
        """
        Gets selected cell.

        Raises:
            CellError: if active selection is not a single cell

        Returns:
            CalcCell: Selected cell
        """
        addr = mCalc.Calc.get_selected_cell_addr(self.calc_doc.component)
        return self.get_cell(addr=addr)

    def get_selected(self) -> mCalcCellRange.CalcCellRange:
        """
        Gets selected cell range.

        Returns:
            CalcCellRange: Selected cell range
        """
        range_obj = self.get_selected_range()
        return mCalcCellRange.CalcCellRange(self, range_obj)

    def get_selected_addr(self) -> CellRangeAddress:
        """
        Gets select cell range addresses

        Raises:
            Exception: if unable to get document model
            MissingInterfaceError: if unable to get interface XCellRangeAddressable

        Returns:
            CellRangeAddress: Cell range addresses.

        See Also:
            - :py:meth:`~.Calc.get_selected_range`
            - :py:meth:`~.Calc.set_selected_addr`
            - :py:meth:`~.Calc.set_selected_range`
            - :py:meth:`~.Calc.get_selected_cell_addr`
        """
        return mCalc.Calc.get_selected_addr(self.calc_doc.component)

    def get_selected_range(self) -> mRngObj.RangeObj:
        """
        Gets select cell range

        Raises:
            Exception: if unable to get document model
            MissingInterfaceError: if unable to get interface XCellRangeAddressable

        Returns:
            RangeObj: Cell range addresses

        See Also:
            - :py:meth:`~.Calc.get_selected_addr`
            - :py:meth:`~.Calc.set_selected_addr`
            - :py:meth:`~.Calc.get_selected_cell_addr`
            - :py:meth:`~.Calc.set_selected_range`
        """
        return mCalc.Calc.get_selected_range(self.calc_doc.component)

    def get_sheet_index(self) -> int:
        """
        Gets index if sheet

        Args:
            sheet (XSpreadsheet | None, optional): Spread sheet. Defaults to active sheet.

        Returns:
            int: _description_
        """
        return mCalc.Calc.get_sheet_index(self.component)

    def get_sheet_name(self, safe_quote: bool = True) -> str:
        """
        Gets the name of a sheet

        Args:
            safe_quote (bool, optional): If True, returns quoted (in single quotes) sheet name if the sheet name is not alphanumeric.
                Defaults to True.

        Raises:
            MissingInterfaceError: If unable to access spreadsheet named interface

        Returns:
            str: Name of sheet
        """
        return mCalc.Calc.get_sheet_name(self.component, safe_quote=safe_quote)

    def select_cells_addr(self, range_val: str | mRngObj.RangeObj) -> CellRangeAddress | None:
        """
        Selects cells in a Spreadsheet.

        Args:
            range_val (str | RangeObj): Range name

        Returns:
            CellRangeAddress | None: Cell range address of the current selection if successful, otherwise ``None``

        See Also:
            - :py:meth:`~.Calc.get_selected_addr`
            - :py:meth:`~.Calc.get_selected_cell_addr`
        """
        return mCalc.Calc.set_selected_addr(doc=self.calc_doc.component, sheet=self.component, range_val=range_val)

    def select_cells_range(self, range_val: str | mRngObj.RangeObj) -> mRngObj.RangeObj | None:
        """
        Selects cells in a Spreadsheet.

        Args:
            range_val (str | RangeObj): Range name

        Returns:
            RangeObj | None: Cell range of the current selection if successful, otherwise ``None``

        See Also:
            - :py:meth:`~.Calc.get_selected_range`
            - :py:meth:`~.Calc.get_selected_cell_range`
        """
        return mCalc.Calc.set_selected_range(doc=self.calc_doc.component, sheet=self.component, range_val=range_val)

    def select_cells_calc_cell_range(self, range_val: str | mRngObj.RangeObj) -> mCalcCellRange.CalcCellRange | None:
        """
        Selects cells in a Spreadsheet.

        Args:
            range_val (str | RangeObj): Range name

        Returns:
            CalcCellRange | None: Cell range of the current selection if successful, otherwise ``None``
        """
        result = self.select_cells_range(range_val)
        if result is None:
            return None
        return mCalcCellRange.CalcCellRange(self, result)

    def set_sheet_name(self, name: str) -> bool:
        """
        Sets the name of a spreadsheet.

        Args:
            name (str): New name for spreadsheet.

        Returns:
            bool: True on success; Otherwise, False
        """
        return mCalc.Calc.set_sheet_name(self.component, name)

    # region set_array()
    @overload
    def set_array(self, *, values: Table, cell_range: XCellRange) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            cell_range (XCellRange): Range in spreadsheet to insert data.

        Returns:
            None:
        """
        ...

    @overload
    def set_array(self, *, values: Table, cell_range: XCellRange, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            cell_range (XCellRange): Range in spreadsheet to insert data.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    def set_array(self, *, values: Table, name: str) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            name (str): Range name such as 'A1:D4' or cell name such as 'B4'.

        Returns:
            None:
        """
        ...

    @overload
    def set_array(self, *, values: Table, name: str, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            name (str): Range name such as 'A1:D4' or cell name such as 'B4'.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    def set_array(self, *, values: Table, range_obj: mRngObj.RangeObj) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            range_obj (RangeObj): Range Object.

        Returns:
            None:
        """
        ...

    @overload
    def set_array(self, *, values: Table, range_obj: mRngObj.RangeObj, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            range_obj (RangeObj): Range Object.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        ...

    @overload
    def set_array(self, *, values: Table, cell_obj: mCellObj.CellObj) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            cell_obj (CellObj): Cell Object

        Returns:
            None:
        """
        ...

    @overload
    def set_array(self, *, values: Table, cell_obj: mCellObj.CellObj, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            cell_obj (CellObj): Cell Object
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    def set_array(self, *, values: Table, addr: CellAddress) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            addr (CellAddress): Address to insert data.

        Returns:
            None:
        """
        ...

    @overload
    def set_array(self, *, values: Table, addr: CellAddress, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            addr (CellAddress): Address to insert data.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    def set_array(self, *, values: Table, col_start: int, row_start: int, col_end: int, row_end: int) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            col_start (int): Zero-base Start Column.
            row_start (int): Zero-base Start Row.
            col_end (int): Zero-base End Column.
            row_end (int): Zero-base End Row.

        Returns:
            None:
        """
        ...

    @overload
    def set_array(
        self, *, values: Table, col_start: int, row_start: int, col_end: int, row_end: int, styles: Sequence[StyleT]
    ) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            col_start (int): Zero-base Start Column.
            row_start (int): Zero-base Start Row.
            col_end (int): Zero-base End Column.
            row_end (int): Zero-base End Row.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    def set_array(self, **kwargs) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            cell_range (XCellRange): Range in spreadsheet to insert data.
            name (str): Range name such as 'A1:D4' or cell name such as 'B4'.
            range_obj (RangeObj): Range Object.
            cell_obj (CellObj): Cell Object.
            addr (CellAddress): Address to insert data.
            col_start (int): Zero-base Start Column.
            row_start (int): Zero-base Start Row.
            col_end (int): Zero-base End Column.
            row_end (int): Zero-base End Row.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        kargs = kwargs.copy()
        sheet_names = {"name", "range_obj", "cell_obj"}

        if kargs.keys() & sheet_names:
            kargs["sheet"] = self.component

        if "addr" in kargs:
            kargs["doc"] = self.calc_doc.component

        mCalc.Calc.set_array(**kargs)

    # endregion set_array()

    # region set_array_range()
    @overload
    def set_array_range(self, *, range_name: str, values: Table) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            range_name (str): Range to insert data such as 'A1:E12'.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.

        Returns:
            None:
        """
        ...

    @overload
    def set_array_range(self, *, range_name: str, values: Table, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            range_name (str): Range to insert data such as 'A1:E12'.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    def set_array_range(self, *, range_obj: mRngObj.RangeObj, values: Table) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            range_obj (RangeObj): Range Object.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.

        Returns:
            None:
        """
        ...

    @overload
    def set_array_range(self, *, range_obj: mRngObj.RangeObj, values: Table, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            range_obj (RangeObj): Range Object.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    def set_array_range(self, **kwargs) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            range_name (str): Range to insert data such as 'A1:E12'.
            range_obj (RangeObj): Range Object.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        mCalc.Calc.set_array_range(self.component, **kwargs)

    # endregion set_array_range()

    # region set_array_cell()
    @overload
    def set_array_cell(self, range_name: str, values: Table) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            range_name (str): Range to insert data such as 'A1:E12'.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
        """
        ...

    @overload
    def set_array_cell(self, range_name: str, values: Table, *, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            range_name (str): Range to insert data such as 'A1:E12'.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.
        """
        ...

    @overload
    def set_array_cell(self, cell_obj: mCellObj.CellObj, values: Table) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            cell_obj (CellObj): Range Object.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
        """
        ...

    @overload
    def set_array_cell(self, cell_obj: mCellObj.CellObj, values: Table, *, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            cell_obj (CellObj): Range Object.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.
        """
        ...

    def set_array_cell(self, *args, **kwargs) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            range_name (str): Range to insert data such as 'A1:E12'.
            cell_obj (CellObj): Range Object.
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        mCalc.Calc.set_array_cell(self.component, *args, **kwargs)

    # endregion set_array_cell()
    @overload
    def set_cell_range_array(self, cell_range: XCellRange, values: Table) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            cell_range (XCellRange): Cell Range
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.

        Returns:
            None:
        """
        ...

    @overload
    def set_cell_range_array(self, cell_range: XCellRange, values: Table, styles: Sequence[StyleT]) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            cell_range (XCellRange): Cell Range
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    def set_cell_range_array(
        self, cell_range: XCellRange, values: Table, styles: Sequence[StyleT] | None = None
    ) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            cell_range (XCellRange): Cell Range
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        if styles:
            mCalc.Calc.set_cell_range_array(cell_range, values, styles)
        else:
            mCalc.Calc.set_cell_range_array(cell_range, values)

    # region set_cell_range_array()

    # region set_col()
    @overload
    def set_col(self, values: Column, cell_name: str) -> None:
        """
        Inserts a column of data into spreadsheet.

        Args:
            values (Column): Column Data.
            cell_name (str): Name of Cell to begin the insert such as 'A1'.
        """
        ...

    @overload
    def set_col(self, values: Column, cell_obj: mCellObj.CellObj) -> None:
        """
        Inserts a column of data into spreadsheet.

        Args:
            values (Column): Column Data.
            cell_obj (CellObj): Cell Object.
        """
        ...

    @overload
    def set_col(self, values: Column, col_start: int, row_start: int) -> None:
        """
        Inserts a column of data into spreadsheet.

        Args:
            values (Column): Column Data.
            col_start (int): Zero-base column index.
            row_start (int): Zero-base row index.
        """
        ...

    def set_col(self, *args, **kwargs) -> None:
        """
        Inserts a column of data into spreadsheet.

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            values (Column): Column Data.
            cell_name (str): Name of Cell to begin the insert such as 'A1'.
            cell_obj (CellObj): Cell Object.
            col_start (int): Zero-base column index.
            row_start (int): Zero-base row index.
        """
        mCalc.Calc.set_col(self.component, *args, **kwargs)

    # endregion set_col()

    def set_col_width(self, width: int | UnitT, idx: int) -> mCalcCellRange.CalcCellRange | None:
        """
        Sets column width. width is in ``mm``, e.g. ``6``

        Args:
            width (int, UnitT): Width in ``mm`` units or :ref:`proto_unit_obj`.
            idx (int): Index of column.

        Raises:
            CancelEventError: If SHEET_COL_WIDTH_SETTING event is canceled.

        Returns:
            CalcCellRange | None: Column cell range that width is applied on or ``None`` if column width <= 0

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_COL_WIDTH_SETTING` :eventref:`src-docs-sheet-event-col-width-setting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_COL_WIDTH_SET` :eventref:`src-docs-sheet-event-col-width-set`

        Note:
            Event args ``index`` is set to ``idx`` value, ``event_data`` is set to ``width`` value (``mm100`` units).
        """
        result = mCalc.Calc.set_col_width(sheet=self.component, width=width, idx=idx)
        if result is None:
            return None
        return mCalcCellRange.CalcCellRange(self, result)

    # endregion set_cell_range_array()

    # region set_row()
    @overload
    def set_row(self, values: Row, cell_name: str) -> None:
        """
        Inserts a row of data into spreadsheet

        Raises:
            MissingInterfaceError: if unable to obtain interface

        Args:
            values (Row): Row Data.
            cell_name (str): Name of Cell to begin the insert such as 'A1'.
        """
        ...

    @overload
    def set_row(self, values: Row, cell_obj: mCellObj.CellObj) -> None:
        """
        Inserts a row of data into spreadsheet

        Raises:
            MissingInterfaceError: if unable to obtain interface

        Args:
            values (Row): Row Data.
            cell_obj (CellObj): Cell Object.
        """
        ...

    @overload
    def set_row(self, values: Row, col_start: int, row_start: int) -> None:
        """
        Inserts a row of data into spreadsheet

        Raises:
            MissingInterfaceError: if unable to obtain interface

        Args:
            values (Row): Row Data.
            col_start (int): Zero-base column index.
            row_start (int): Zero-base row index.
        """
        ...

    def set_row(self, *args, **kwargs) -> None:
        """
        Inserts a row of data into spreadsheet

        Raises:
            MissingInterfaceError: if unable to obtain interface

        Args:
            sheet (XSpreadsheet): Spreadsheet.
            values (Row): Row Data.
            cell_obj (CellObj): Cell Object.
            cell_name (str): Name of Cell to begin the insert such as 'A1'.
            col_start (int): Zero-base column index.
            row_start (int): Zero-base row index.
        """
        mCalc.Calc.set_row(self.component, *args, **kwargs)

    # endregion set_row()

    def set_row_height(self, height: int | UnitT, idx: int) -> mCalcCellRange.CalcCellRange | None:
        """
        Sets column width. height is in ``mm``, e.g. 6

        Args:
            height (int, UnitT): Width in ``mm`` units or :ref:`proto_unit_obj`.
            idx (int): Index of Row

        Raises:
            CancelEventError: If SHEET_ROW_HEIGHT_SETTING event is canceled.

        Returns:
            CalcCellRange | None: Row cell range that height is applied on or None if height <= 0

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_HEIGHT_SETTING` :eventref:`src-docs-sheet-event-row-height-setting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_HEIGHT_SET` :eventref:`src-docs-sheet-event-row-height-set`

        Note:
            Event args ``index`` is set to ``idx`` value, ``event_data`` is set to ``height`` value (``mm100`` units).
        """
        result = mCalc.Calc.set_row_height(sheet=self.component, height=height, idx=idx)
        if result is None:
            return None
        return mCalcCellRange.CalcCellRange(self, result)

    # region set_value()
    @overload
    def set_val(self, *, value: object, cell: XCell) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell.
            cell (XCell): Cell to assign value.

        Returns:
            None:
        """
        ...

    @overload
    def set_val(self, *, value: object, cell: XCell, styles: Sequence[StyleT]) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell.
            cell (XCell): Cell to assign value.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell.

        Returns:
            None:
        """
        ...

    @overload
    def set_val(self, *, value: object, cell_name: str) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell.
            cell_name (str): Name of cell to set value of such as 'B4'.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell.

        Returns:
            None:
        """
        ...

    @overload
    def set_val(self, *, value: object, cell_name: str, styles: Sequence[StyleT]) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell.
            cell_name (str): Name of cell to set value of such as 'B4'.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell.

        Returns:
            None:
        """
        ...

    @overload
    def set_val(self, *, value: object, cell_obj: mCellObj.CellObj) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell.
            cell_obj (CellObj): Cell Object.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        ...

    @overload
    def set_val(self, *, value: object, cell_obj: mCellObj.CellObj, styles: Sequence[StyleT]) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell.
            cell_obj (CellObj): Cell Object.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell.

        Returns:
            None:
        """
        ...

    @overload
    def set_val(self, *, value: object, col: int, row: int) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell.
            col (int): Cell column as zero-based integer.
            row (int): Cell row as zero-based integer.

        Returns:
            None:
        """
        ...

    @overload
    def set_val(self, *, value: object, col: int, row: int, styles: Sequence[StyleT]) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell.
            col (int): Cell column as zero-based integer.
            row (int): Cell row as zero-based integer.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell.

        Returns:
            None:
        """
        ...

    def set_val(self, **kwargs) -> None:
        """
        Sets the value of a cell

        Args:
            value (object): Value for cell.
            cell (XCell): Cell to assign value.
            cell_name (str): Name of cell to set value of such as 'B4'.
            cell_obj (CellObj): Cell Object.
            col (int): Cell column as zero-based integer.
            row (int): Cell row as zero-based integer.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        sheet_names = {"cell_name", "cell_obj", "col"}
        if kwargs.keys() & sheet_names:
            kwargs["sheet"] = self.component
        mCalc.Calc.set_val(**kwargs)

    # endregion set_value()

    def split_window(self, cell_name: str) -> None:
        """
        Splits window

        Args:
            cell_name (str): Cell to preform split on. e.g. 'C4'

        Returns:
            None:

        See Also:
            :ref:`ch23_splitting_panes`
        """
        mCalc.Calc.split_window(doc=self.calc_doc.component, cell_name=cell_name)

    # region unmerge_cells()
    @overload
    def unmerge_cells(self, *, cell_range: XCellRange) -> None:
        """
        Removes merging from a range of cells

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            None:
        """
        ...

    @overload
    def unmerge_cells(self, *, range_name: str) -> None:
        """
        Removes merging from a range of cells

        Args:
            range_name (str): Range Name such as ``A1:D5``.

        Returns:
            None:
        """
        ...

    @overload
    def unmerge_cells(self, *, range_obj: mRngObj.RangeObj) -> None:
        """
        Removes merging from a range of cells

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            None:
        """
        ...

    @overload
    def unmerge_cells(self, *, cr_addr: CellRangeAddress) -> None:
        """
        Removes merging from a range of cells

        Args:
            cr_addr (CellRangeAddress): Cell range Address.

        Returns:
            None:
        """
        ...

    @overload
    def unmerge_cells(self, *, col_start: int, row_start: int, col_end: int, row_end: int) -> None:
        """
        Removes merging from a range of cells

        Args:
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Returns:
            None:
        """
        ...

    def unmerge_cells(self, **kwargs) -> None:
        """
        Removes merging from a range of cells

        Args:
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            cr_addr (CellRangeAddress): Cell range Address.
            cell_range (XCellRange): Cell Range.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Returns:
            None:

        See Also:
            - :py:meth:`.Calc.merge_cells`
            - :py:meth:`.Calc.is_merged_cells`
        """
        sheet_names = {"range_name", "range_obj", "cr_addr"}
        if kwargs.keys() & sheet_names:
            kwargs["sheet"] = self.component
        mCalc.Calc.unmerge_cells(**kwargs)

    # endregion unmerge_cells()

    # region    goto_cell()
    @overload
    def goto_cell(self, cell_name: str) -> mCalcCell.CalcCell:
        ...

    @overload
    def goto_cell(self, cell_obj: mCellObj.CellObj) -> mCalcCell.CalcCell:
        ...

    def goto_cell(self, *args, **kwargs) -> mCalcCell.CalcCell:
        """
        Go to a cell

        Args:
            cell_name (str): Cell Name such as 'B4'.
            doc (XSpreadsheetDocument): Spreadsheet Document.
            frame (XFrame): Spreadsheet frame.

        Returns:
            CalcCell: Cell Object.

        Attention:
            :py:meth:`~.utils.lo.Lo.dispatch_cmd` method is called along with any of its events.

            Dispatch command is ``GoToCell``.
        """

        def go(obj) -> mCellObj.CellObj:
            frame = self.calc_doc.get_controller().getFrame()
            s = str(obj)
            props = mProps.Props.make_props(ToPoint=str(obj))
            mLo.Lo.dispatch_cmd(cmd="GoToCell", props=props, frame=frame)
            return mCellObj.CellObj.from_cell(cell_val=s)

        kargs_len = len(kwargs)
        count = len(args) + kargs_len

        if count == 0:
            raise TypeError("goto_cell() missing 1 required positional argument: 'cell_name' or 'cell_obj'")
        if "cell_name" in kwargs:
            cell_obj = go(kwargs["cell_name"])
            return self.get_cell(cell_obj=cell_obj)
        if "cell_obj" in kwargs:
            cell_obj = go(kwargs["cell_obj"])
            return self.get_cell(cell_obj=cell_obj)
        if kwargs:
            raise TypeError("goto_cell() got an unexpected keyword argument")
        if count != 1:
            raise TypeError("goto_cell() got an invalid number of arguments")
        arg = args[0]
        cell_obj = go(arg)
        return self.get_cell(cell_obj=cell_obj)

    # endregion    goto_cell()

    def delete_column(self, idx: int, count: int = 1) -> bool:
        """
        Delete a column from a spreadsheet.

        Args:
            idx (int): Zero base of index of column to delete.
            count (int): Number of columns to delete.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_COL_DELETING` :eventref:`src-docs-sheet-event-col-deleting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_COL_DELETED` :eventref:`src-docs-sheet-event-col-deleted`

        Returns:
            bool: ``True`` if column is deleted; Otherwise, ``False``.
        """
        return mCalc.Calc.delete_column(sheet=self.component, idx=idx, count=count)

    def delete_row(self, idx: int, count: int = 1) -> bool:
        """
        Deletes a row from spreadsheet.

        Args:
            idx (int): Zero based index of row to delete.
            count (int): Number of rows to delete.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_DELETING` :eventref:`src-docs-sheet-event-row-deleting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_DELETED` :eventref:`src-docs-sheet-event-row-deleted`

        Returns:
            bool: ``True`` if row is deleted; Otherwise, ``False``.
        """
        return mCalc.Calc.delete_row(sheet=self.component, idx=idx, count=count)

    def deselect_cells(self) -> None:
        """
        Deselects cells in a Spreadsheet.

        Returns:
            None:
        """
        mCalc.Calc.set_selected_addr(doc=self.calc_doc.component, sheet=self.component)

    def dispatch_recalculate(self) -> None:
        """
        Dispatches recalculate command to the current sheet.

        Also useful when needing to refresh a chart.

        Returns:
            None:
        """
        mCalc.Calc.dispatch_recalculate()

    def extract_col(self, vals: Table, col_idx: int) -> List[Any]:
        """
        Extract column data and returns as a list

        Args:
            vals (Table): 2-d table of data
            col_idx (int): column index to extract

        Returns:
            List[Any]: Column data if found; Otherwise, empty list.
        """
        return mCalc.Calc.extract_col(vals=vals, col_idx=col_idx)

    def extract_row(self, vals: Table, row_idx: int) -> Row:
        """
        Extracts a row from a table

        Args:
            vals (Table): Table of data
            row_idx (int): Row index to extract

        Raises:
            IndexError: If row_idx is out of range.

        Returns:
            Row: Row of data
        """
        return mCalc.Calc.extract_row(vals=vals, row_idx=row_idx)

    def insert_row(self, idx: int, count: int = 1) -> bool:
        """
        Inserts a row into spreadsheet.

        Args:
            idx (int): Zero based index of row to insert.
            count (int): Number of rows to insert.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_INSERTING` :eventref:`src-docs-sheet-event-row-inserting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ROW_INSERTED` :eventref:`src-docs-sheet-event-row-inserted`

        Returns:
            bool: ``True`` if row is inserted; Otherwise, ``False``.
        """
        return mCalc.Calc.insert_row(sheet=self.component, idx=idx, count=count)

    def insert_column(self, idx: int, count: int = 1) -> bool:
        """
        Inserts a column into spreadsheet.

        Args:
            idx (int): Zero based index of column to insert.
            count (int): Number of columns to insert.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_COL_INSERTING` :eventref:`src-docs-sheet-event-col-inserting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_COL_INSERTED` :eventref:`src-docs-sheet-event-col-inserted`

        Returns:
            bool: ``True`` if column is inserted; Otherwise, ``False``.
        """
        return mCalc.Calc.insert_column(sheet=self.component, idx=idx, count=count)

    def insert_scenario(self, range_name: str | mRngObj.RangeObj, vals: Table, name: str, comment: str) -> XScenario:
        """
        Insert a scenario into sheet

        Args:
            range_name (str | RangeObj): Range name
            vals (Table): 2d array of values
            name (str): Scenario name
            comment (str): Scenario description

        Returns:
            XScenario: the newly created scenario

        Note:
            A LibreOffice Calc scenario is a set of cell values that can be used within your calculations.
            You assign a name to every scenario on your sheet. Define several scenarios on the same sheet,
            each with some different values in the cells. Then you can easily switch the sets of cell values
            by their name and immediately observe the results. Scenarios are a tool to test out "what-if" questions.

        See Also:
            `Using Scenarios <https://help.libreoffice.org/latest/en-US/text/scalc/guide/scenario.html>`_
        """
        return mCalc.Calc.insert_scenario(
            sheet=self.component, range_name=range_name, vals=vals, name=name, comment=comment
        )

    # region clear_cells()
    @overload
    def clear_cells(self, cell_range: XCellRange) -> bool:
        """
        Clears the specified contents of the cell range.

        Cell range of types ``VALUE``, ``DATETIME`` and ``STRING`` are cleared.

        Args:
            cell_range (XCellRange): Cell range.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @overload
    def clear_cells(self, cell_range: XCellRange, cell_flags: CellFlagsEnum) -> bool:
        """
        Clears the specified contents of the cell range.


        Args:
            cell_range (XCellRange): Cell range.
            cell_flags (CellFlagsEnum): Flags that determine what to clear.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @overload
    def clear_cells(self, range_name: str) -> bool:
        """
        Clears the specified contents of the cell range.

        Cell range of types ``VALUE``, ``DATETIME`` and ``STRING`` are cleared.

        Args:
            range_name (str): Range name such as ``A1:G3``.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @overload
    def clear_cells(self, range_name: str, cell_flags: CellFlagsEnum) -> bool:
        """
        Clears the specified contents of the cell range.

        Args:
            range_name (str): Range name such as ``A1:G3``.
            cell_flags (CellFlagsEnum): Flags that determine what to clear.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @overload
    def clear_cells(self, range_val: mRngObj.RangeObj, cell_flags: CellFlagsEnum) -> bool:
        """
        Clears the specified contents of the cell range.

        Args:
            range_val (RangeObj): Range object.
            cell_flags (CellFlagsEnum): Flags that determine what to clear.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    @overload
    def clear_cells(self, cr_addr: CellRangeAddress) -> bool:
        """
        Clears the specified contents of the cell range.

        Cell range of types ``VALUE``, ``DATETIME`` and ``STRING`` are cleared.

        Args:
            cr_addr (CellRangeAddress): Cell Range Address.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``/
        """
        ...

    @overload
    def clear_cells(self, cr_addr: CellRangeAddress, cell_flags: CellFlagsEnum) -> bool:
        """
        Clears the specified contents of the cell range.

        Args:
            cr_addr (CellRangeAddress): Cell Range Address.
            cell_flags (CellFlagsEnum): Flags that determine what to clear.

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.
        """
        ...

    def clear_cells(self, *args, **kwargs) -> bool:
        """
        Clears the specified contents of the cell range.

        If cell_flags is not specified then
        cell range of types ``VALUE``, ``DATETIME`` and ``STRING`` are cleared.

        Args:
            cell_range (XCellRange): Cell range.
            range_name (str): Range name such as ``A1:G3``.
            range_val (RangeObj): Range object.
            cr_addr (CellRangeAddress): Cell Range Address.
            cell_flags (CellFlagsEnum): Flags that determine what to clear.

        Raises:
            MissingInterfaceError: If XSheetOperation interface cannot be obtained.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_CLEARING` :eventref:`src-docs-cell-event-clearing`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_CLEARED` :eventref:`src-docs-cell-event-cleared`

        Returns:
            bool: ``True`` if cells are cleared; Otherwise, ``False``.

        Note:
            Events arg for this method have a ``cell`` type of ``XCellRange``.

            Events arg ``event_data`` is a dictionary containing ``cell_flags``.
        """
        return mCalc.Calc.clear_cells(self.component, *args, **kwargs)

    # endregion clear_cells()

    # region delete_cells()
    @overload
    def delete_cells(self, cell_range: XCellRange, is_shift_left: bool) -> bool:
        """
        Deletes Cells in a spreadsheet.

        Args:
            cell_range (XCellRange): Cell range to insert.
            is_shift_left (bool): If ``True`` then cell are shifted left; Otherwise, cells are shifted up.

        Returns:
            bool: ``True`` if cells are deleted; Otherwise, ``False``.
        """
        ...

    @overload
    def delete_cells(self, cr_addr: CellRangeAddress, is_shift_left: bool) -> bool:
        """
        Deletes Cells in a spreadsheet.

        Args:
            cr_addr (CellRangeAddress): Cell range Address.
            is_shift_left (bool): If ``True`` then cell are shifted left; Otherwise, cells are shifted up.

        Returns:
            bool: ``True`` if cells are deleted; Otherwise, ``False``.
        """
        ...

    @overload
    def delete_cells(self, range_name: str, is_shift_left: bool) -> bool:
        """
        Deletes Cells in a spreadsheet.

        Args:
            range_name (str): Range Name such as 'A1:D5'.
            is_shift_left (bool): If ``True`` then cell are shifted left; Otherwise, cells are shifted up.

        Returns:
            bool: ``True`` if cells are deleted; Otherwise, ``False``.
        """
        ...

    @overload
    def delete_cells(self, range_obj: mRngObj.RangeObj, is_shift_left: bool) -> bool:
        """
        Deletes Cells in a spreadsheet.

        Args:
            range_obj (RangeObj): Range Object.
            is_shift_left (bool): If ``True`` then cell are shifted left; Otherwise, cells are shifted up.

        Returns:
            bool: ``True`` if cells are deleted; Otherwise, ``False``.
        """
        ...

    @overload
    def delete_cells(self, col_start: int, row_start: int, col_end: int, row_end: int, is_shift_left: bool) -> bool:
        """
        Deletes Cells in a spreadsheet.

        Args:
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            is_shift_left (bool): If ``True`` then cell are shifted left; Otherwise, cells are shifted up.

        Returns:
            bool: ``True`` if cells are deleted; Otherwise, ``False``.
        """
        ...

    def delete_cells(self, *args, **kwargs) -> bool:
        """
        Deletes Cells in a spreadsheet.

        Args:
            cell_range (XCellRange): Cell range to insert.
            cr_addr (CellRangeAddress): Cell range Address.
            range_name (str): Range Name such as 'A1:D5'.
            range_obj (RangeObj): Range Object.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            is_shift_left (bool): If ``True`` then cell are shifted left; Otherwise, cells are shifted up.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_DELETING` :eventref:`src-docs-cell-event-deleting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_DELETED` :eventref:`src-docs-cell-event-deleted`

        Returns:
            bool: ``True`` if cells are deleted; Otherwise, ``False``.

        Note:
            Events args for this method have a ``cell`` type of ``XCellRange``

            Event args ``event_data`` is a dictionary containing ``is_shift_right``.
        """
        return mCalc.Calc.delete_cells(self.component, *args, **kwargs)

    # endregion delete_cells()

    # region insert_cells()
    @overload
    def insert_cells(self, cell_range: XCellRange, is_shift_right: bool) -> bool:
        """
        Inserts Cells into a spreadsheet.

        Args:
            cell_range (XCellRange): Cell range to insert.
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.

        Returns:
            bool: ``True`` if cells are inserted; Otherwise, ``False``.
        """
        ...

    @overload
    def insert_cells(self, cr_addr: CellRangeAddress, is_shift_right: bool) -> bool:
        """
        Inserts Cells into a spreadsheet.

        Args:
            cr_addr (CellRangeAddress): Cell range Address.
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.

        Returns:
            bool: ``True`` if cells are inserted; Otherwise, ``False``.
        """
        ...

    @overload
    def insert_cells(self, range_name: str, is_shift_right: bool) -> bool:
        """
        Inserts Cells into a spreadsheet.

        Args:
            range_name (str): Range Name such as 'A1:D5'.
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.

        Returns:
            bool: ``True`` if cells are inserted; Otherwise, ``False``.
        """
        ...

    @overload
    def insert_cells(self, range_obj: mRngObj.RangeObj, is_shift_right: bool) -> bool:
        """
        Inserts Cells into a spreadsheet.

        Args:
            range_obj (RangeObj): Range Object.
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.

        Returns:
            bool: ``True`` if cells are inserted; Otherwise, ``False``.
        """
        ...

    @overload
    def insert_cells(self, col_start: int, row_start: int, col_end: int, row_end: int, is_shift_right: bool) -> bool:
        """
        Inserts Cells into a spreadsheet.

        Args:
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.

        Returns:
            bool: ``True`` if cells are inserted; Otherwise, ``False``.
        """
        ...

    def insert_cells(self, *args, **kwargs) -> bool:
        """
        Inserts Cells into a spreadsheet.

        Args:
            cell_range (XCellRange): Cell range to insert.
            cr_addr (CellRangeAddress): Cell range Address.
            range_name (str): Range Name such as 'A1:D5'.
            range_obj (RangeObj): Range Object.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            is_shift_right (bool): If True then cell are inserted to the right; Otherwise, inserted down.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_INSERTING` :eventref:`src-docs-cell-event-inserting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_INSERTED` :eventref:`src-docs-cell-event-inserted`

        Returns:
            bool: ``True`` if cells are inserted; Otherwise, ``False``.

        Note:
            Events args for this method have a ``cell`` type of ``XCellRange``

            Event args ``event_data`` is a dictionary containing ``is_shift_right``.
        """
        return mCalc.Calc.insert_cells(self.component, *args, **kwargs)

    # endregion insert_cells()

    def find_all(self, srch: XSearchable, sd: XSearchDescriptor) -> List[mCalcCellRange.CalcCellRange] | None:
        """
        Searches spreadsheet and returns a list of Cell Ranges that match search criteria

        Args:
            srch (XSearchable): Searchable object
            sd (XSearchDescriptor): Search description

        Returns:
            List[XCellRange] | None: A list of cell ranges on success; Otherwise, None


        .. collapse:: Example

            .. code-block:: python

                from ooodev.utils.lo import Lo
                from ooodev.office.calc import Calc
                from com.sun.star.util import XSearchable

                doc = Calc.create_doc(loader)
                sheet = Calc.get_sheet(doc=doc, index=0)
                Calc.set_val(value='test', sheet=sheet, cell_name="A1")
                Calc.set_val(value='test', sheet=sheet, cell_name="C3")
                srch = Lo.qi(XSearchable, sheet)
                sd = srch.createSearchDescriptor()
                sd.setSearchString('test')
                results = Calc.find_all(srch=srch, sd=sd)
                assert len(results) == 2

        See Also:
            `LibreOffice API SearchDescriptor <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1util_1_1SearchDescriptor.html>`_
        """
        found = mCalc.Calc.find_all(srch=srch, sd=sd)
        if not found:
            return None
        return [mCalcCellRange.CalcCellRange(self, x) for x in found]

    # region find_function()
    @overload
    def find_function(self, func_nm: str) -> Tuple[PropertyValue] | None:
        """
        Finds a function

        Args:
            func_nm (str): function name

        Returns:
            Tuple[PropertyValue] | None: Function properties as tuple on success; Otherwise, None
        """
        ...

    @overload
    def find_function(self, idx: int) -> Tuple[PropertyValue] | None:
        """
        Finds a function

        Args:
            idx (int): Index of function

        Returns:
            Tuple[PropertyValue] | None: Function properties as tuple on success; Otherwise, None
        """
        ...

    def find_function(self, *args, **kwargs) -> Tuple[PropertyValue, ...] | None:
        """
        Finds a function

        Args:
            func_nm (str): function name
            idx (int): Index of function

        Returns:
            Tuple[PropertyValue, ...] | None: Function properties as tuple on success; Otherwise, None
        """
        return mCalc.Calc.find_function(*args, **kwargs)

    # endregion find_function()

    def find_used_cursor(self, cursor: XSheetCellCursor) -> mCalcCellRange.CalcCellRange:
        """
        Find used cursor

        Args:
            cursor (CalcCellRange): Sheet Cursor

        Raises:
            MissingInterfaceError: if unable to find interface

        Returns:
            XCellRange: Cell range
        """
        found = mCalc.Calc.find_used_cursor(cursor=cursor)
        return mCalcCellRange.CalcCellRange(self, found)

    # region find_used()
    @overload
    def find_used(
        self,
    ) -> mCalcCellRange.CalcCellRange:
        """
        Find used range

        Returns:
            CalcCellRange: Cell range
        """
        ...

    @overload
    def find_used(self, range_name: str) -> mCalcCellRange.CalcCellRange:
        """
        Find used range

        Args:
            range_name (str): Range Name such as 'A1:D5'

        Returns:
            CalcCellRange: Cell range
        """
        ...

    @overload
    def find_used(self, range_obj: mRngObj.RangeObj) -> mCalcCellRange.CalcCellRange:
        """
        Find used range

        Args:
            range_obj (RangeObj): Range Object

        Returns:
            CalcCellRange: Cell range
        """
        ...

    @overload
    def find_used(self, cr_addr: CellRangeAddress) -> mCalcCellRange.CalcCellRange:
        """
        Find used range

        Args:
            cr_addr (CellRangeAddress): Cell range Address

        Returns:
            CalcCellRange: Cell range
        """
        ...

    def find_used(self, *args, **kwargs) -> mCalcCellRange.CalcCellRange:
        """
        Find used range

        Args:
            range_name (str): Range Name such as 'A1:D5'
            range_obj (RangeObj): Range Object
            cr_addr (CellRangeAddress): Cell range Address

        Returns:
            CalcCellRange: Cell range

        See Also:
            - :ref:`ch20_finding_with_cursors`
        """
        found = mCalc.Calc.find_used_range(self.component, *args, **kwargs)
        return mCalcCellRange.CalcCellRange(self, found)

    # endregion find_used()

    # region find_used_range_obj()
    @overload
    def find_used_range_obj(self) -> mRngObj.RangeObj:
        """
        Find used range

        Returns:
            RangeObj: Range object
        """
        ...

    @overload
    def find_used_range_obj(self, range_name: str) -> mRngObj.RangeObj:
        """
        Find used range

        Args:
            range_name (str): Range Name such as 'A1:D5'

        Returns:
            RangeObj: Range object
        """
        ...

    @overload
    def find_used_range_obj(self, range_obj: mRngObj.RangeObj) -> mRngObj.RangeObj:
        """
        Find used range

        Args:
            range_obj (RangeObj): Range Object

        Returns:
            RangeObj: Range object
        """
        ...

    @overload
    def find_used_range_obj(self, cr_addr: CellRangeAddress) -> mRngObj.RangeObj:
        """
        Find used range

        Args:
            cr_addr (CellRangeAddress): Cell range Address

        Returns:
            RangeObj: Range object
        """
        ...

    def find_used_range_obj(self, *args, **kwargs) -> mRngObj.RangeObj:
        """
        Find used range

        Args:
            range_name (str): Range Name such as 'A1:D5'
            range_obj (RangeObj): Range Object
            cr_addr (CellRangeAddress): Cell range Address

        Returns:
            RangeObj: Range object
        """
        return mCalc.Calc.find_used_range_obj(self.component, *args, **kwargs)

    # endregion find_used_range_obj()

    # region is_merged_cells()
    @overload
    def is_merged_cells(self, *, cell_range: XCellRange) -> bool:
        """
        Gets is a range of cells is merged.

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        ...

    @overload
    def is_merged_cells(self, *, range_name: str) -> bool:
        """
        Gets is a range of cells is merged.

        Args:
            range_name (str): Range Name such as ``A1:D5``.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        ...

    @overload
    def is_merged_cells(self, *, range_obj: mRngObj.RangeObj) -> bool:
        """
        Gets is a range of cells is merged.

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        ...

    @overload
    def is_merged_cells(self, *, cr_addr: CellRangeAddress) -> bool:
        """
        Gets is a range of cells is merged.

        Args:
            cr_addr (CellRangeAddress): Cell range Address.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        ...

    @overload
    def is_merged_cells(self, *, col_start: int, row_start: int, col_end: int, row_end: int) -> bool:
        """
        Gets is a range of cells is merged.

        Args:
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        ...

    def is_merged_cells(self, **kwargs) -> bool:
        """
        Gets is a range of cells is merged.

        Args:
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            cr_addr (CellRangeAddress): Cell range Address.
            cell_range (XCellRange): Cell Range.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``

        See Also:
            - :py:meth:`.Calc.merge_cells`
            - :py:meth:`.Calc.unmerge_cells`
        """
        sheet_names = {"range_name", "range_obj", "cr_addr"}
        if kwargs.keys() & sheet_names:
            kwargs["sheet"] = self.component
        return mCalc.Calc.is_merged_cells(**kwargs)

    # endregion is_merged_cells()

    def is_sheet_protected(self) -> bool:
        """
        Gets whether a sheet is protected

        Returns:
            bool: True if protected; Otherwise, False

        See Also:
            - :py:meth:`~.calc.Calc.protect_sheet`
            - :py:meth:`~.calc.Calc.unprotect_sheet`
            - :ref:`help_calc_format_direct_cell_cell_protection`

        .. versionadded:: 0.10.0
        """
        pro = mLo.Lo.qi(XProtectable, self.component, True)
        return pro.isProtected()

    # region merge_cells()
    @overload
    def merge_cells(self, *, cell_range: XCellRange) -> None:
        """
        Merges a range of cells

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            None:
        """
        ...

    @overload
    def merge_cells(self, *, cell_range: XCellRange, center: bool) -> None:
        """
        Merges a range of cells

        Args:
            cell_range (XCellRange): Cell Range.
            center (bool): Determines if the merge will be a merge and center. Default ``False``.

        Returns:
            None:
        """
        ...

    @overload
    def merge_cells(self, *, range_name: str) -> None:
        """
        Merges a range of cells

        Args:
            range_name (str): Range Name such as ``A1:D5``.

        Returns:
            None:
        """
        ...

    @overload
    def merge_cells(self, *, range_name: str, center: bool) -> None:
        """
        Merges a range of cells

        Args:
            range_name (str): Range Name such as ``A1:D5``.
            center (bool): Determines if the merge will be a merge and center. Default ``False``.

        Returns:
            None:
        """
        ...

    @overload
    def merge_cells(self, *, range_obj: mRngObj.RangeObj) -> None:
        """
        Merges a range of cells

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            None:
        """
        ...

    @overload
    def merge_cells(self, *, range_obj: mRngObj.RangeObj, center: bool) -> None:
        """
        Merges a range of cells

        Args:
            range_obj (RangeObj): Range Object.
            center (bool): Determines if the merge will be a merge and center. Default ``False``.

        Returns:
            None:
        """
        ...

    @overload
    def merge_cells(self, *, cr_addr: CellRangeAddress) -> None:
        """
        Merges a range of cells

        Args:
            cr_addr (CellRangeAddress): Cell range Address.

        Returns:
            None:
        """
        ...

    @overload
    def merge_cells(self, *, cr_addr: CellRangeAddress, center: bool) -> None:
        """
        Merges a range of cells

        Args:
            cr_addr (CellRangeAddress): Cell range Address.
            center (bool): Determines if the merge will be a merge and center. Default ``False``.

        Returns:
            None:
        """
        ...

    @overload
    def merge_cells(self, *, col_start: int, row_start: int, col_end: int, row_end: int, center: bool) -> None:
        """
        Merges a range of cells

        Args:
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            center (bool): Determines if the merge will be a merge and center. Default ``False``.

        Returns:
            None:
        """
        ...

    def merge_cells(self, **kwargs) -> None:
        """
        Merges a range of cells

        Args:
            center (bool): Determines if the merge will be a merge and center. Default ``False``.
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            cr_addr (CellRangeAddress): Cell range Address.
            cell_range (XCellRange): Cell Range.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Returns:
            None:

        See Also:
            - :py:meth:`.Calc.unmerge_cells`
            - :py:meth:`.Calc.is_merged_cells`
        """
        sheet_names = {"range_name", "range_obj", "cr_addr"}
        if kwargs.keys() & sheet_names:
            kwargs["sheet"] = self.component
        mCalc.Calc.merge_cells(**kwargs)

    # endregion merge_cells()

    def protect_sheet(self, password: str) -> bool:
        """
        Protects a Spreadsheet

        Args:
            password (str): Password to protect sheet with.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``

        See Also:
            - :py:meth:`~.calc.Calc.unprotect_sheet`
            - :py:meth:`~.calc.Calc.is_sheet_protected`
            - :ref:`help_calc_format_direct_cell_cell_protection`
        """
        return mCalc.Calc.protect_sheet(self.component, password)

    def unprotect_sheet(self, password: str) -> bool:
        """
        Unprotect a Spreadsheet.

        If sheet is not protected, this method will still return ``True``.

        If incorrect password is provided, this method will return ``False``.

        Args:
            password (str): Password to unprotect sheet with.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``

        See Also:
            - :py:meth:`~.calc.Calc.protect_sheet`
            - :py:meth:`~.calc.Calc.is_sheet_protected`
            - :ref:`help_calc_format_direct_cell_cell_protection`

        .. versionadded:: 0.10.0
        """
        return mCalc.Calc.unprotect_sheet(self.component, password)

    def set_active(self) -> None:
        """
        Sets the current sheet as active in the document.

        Returns:
            None:
        """
        self.calc_doc.set_active_sheet(self.component)

    def create_cursor(self) -> mCalcCellCursor.CalcCellCursor:
        """
        Creates a cell cursor including the whole spreadsheet.

        Returns:
            _type_: _description_
        """
        cursor = self.component.createCursor()
        return mCalcCellCursor.CalcCellCursor(self, cursor)

    # region create_cursor_by_range()
    @overload
    def create_cursor_by_range(self, *, range_name: str) -> mCalcCellCursor.CalcCellCursor:
        """
        Gets a cell cursor.

        Args:
            range_name (str): Range Name such as ``A1:D5``.

        Raises:
            Exception: if unable to access cell range.

        Returns:
            CalcCellCursor: Cell Cursor.
        """
        ...

    @overload
    def create_cursor_by_range(self, *, range_obj: mRngObj.RangeObj) -> mCalcCellCursor.CalcCellCursor:
        """
        Gets a cell cursor.

        Args:
            range_obj (RangeObj): Range Object.

        Raises:
            Exception: if unable to access cell range.

        Returns:
            CalcCellCursor: Cell Cursor.
        """
        ...

    @overload
    def create_cursor_by_range(self, *, cell_obj: mCellObj.CellObj) -> mCalcCellCursor.CalcCellCursor:
        """
        Gets a cell cursor.

        Args:
            cell_obj (CellObj): Cell Object.

        Raises:
            Exception: if unable to access cell range.

        Returns:
            CalcCellCursor: Cell Cursor.
        """
        ...

    @overload
    def create_cursor_by_range(self, *, cr_addr: CellRangeAddress) -> mCalcCellCursor.CalcCellCursor:
        """
        Gets a cell cursor.

        Args:
            cr_addr (CellRangeAddress): Cell range Address.

        Raises:
            Exception: if unable to access cell range.

        Returns:
            CalcCellCursor: Cell Cursor.
        """
        ...

    @overload
    def create_cursor_by_range(self, *, cell_range: XCellRange) -> mCalcCellCursor.CalcCellCursor:
        """
        Gets a cell cursor.

        Args:
            cell_range (XCellRange): Cell Range. If passed in then the same instance is returned.

        Raises:
            Exception: if unable to access cell range.

        Returns:
            CalcCellCursor: Cell Cursor.
        """
        ...

    @overload
    def create_cursor_by_range(
        self, *, col_start: int, row_start: int, col_end: int, row_end: int
    ) -> mCalcCellCursor.CalcCellCursor:
        """
        Gets a cell cursor.

        Args:
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Raises:
            Exception: if unable to access cell range.

        Returns:
            CalcCellCursor: Cell Cursor.
        """
        ...

    def create_cursor_by_range(self, **kwargs) -> mCalcCellCursor.CalcCellCursor:
        """
        Creates a cell cursor to travel in the given range context.

        Args:
            range_name (str): Range Name such as ``A1:D5``.
            range_obj (RangeObj): Range Object.
            cell_obj (CellObj): Cell Object.
            cr_addr (CellRangeAddress): Cell range Address.
            cell_range (XCellRange): Cell Range. If passed in then the same instance is returned.
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.

        Returns:
            CalcCellCursor: Cell cursor
        """
        sheet_names = {"range_name", "range_obj", "cell_obj", "cr_addr", "col_start"}
        if kwargs.keys() & sheet_names:
            kwargs["sheet"] = self.component
        cell_range = mCalc.Calc.get_cell_range(**kwargs)
        sheet_cell_range = mLo.Lo.qi(XSheetCellRange, cell_range, True)
        cursor = self.component.createCursorByRange(sheet_cell_range)
        return mCalcCellCursor.CalcCellCursor(self, cursor)

    # endregion create_cursor_by_range()

    # region Properties
    @property
    def calc_doc(self) -> CalcDoc:
        """
        Returns:
            CalcDoc: Calc doc
        """
        return self.__owner

    @property
    def sheet_name(self) -> str:
        """
        Gets the sheet Name.

        Returns:
            str: Sheet name
        """
        return self.component.getName()

    @property
    def sheet_index(self) -> int:
        """
        Gets the sheet index.

        Returns:
            int: Sheet index
        """
        return self.component.getRangeAddress().Sheet

    # endregion Properties
