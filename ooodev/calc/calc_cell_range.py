from __future__ import annotations
from typing import Any, List, Sequence, TYPE_CHECKING
import uno


from com.sun.star.sheet import XCellSeries
from com.sun.star.table import XCellRange
from ooo.dyn.sheet.cell_flags import CellFlagsEnum as CellFlagsEnum


if TYPE_CHECKING:
    from ooo.dyn.table.cell_range_address import CellRangeAddress
    from .calc_sheet import CalcSheet
    from ooodev.utils.data_type.range_obj import RangeObj
    from ooodev.utils.type_var import Table, TupleArray, FloatTable, Row
    from ooodev.utils.data_type.size import Size
    from . import calc_cell_cursor as mCalcCellCursor
else:
    CellRangeAddress = object

from ooodev.proto.style_obj import StyleT
from ooodev.office import calc as mCalc
from ooodev.adapter.sheet.sheet_cell_range_comp import SheetCellRangeComp
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils import lo as mLo
from . import calc_cell as mCalcCell


class CalcCellRange(SheetCellRangeComp, QiPartial, PropPartial):
    """Represents a calc cell range."""

    def __init__(self, owner: CalcSheet, rng: Any) -> None:
        """
        Constructor

        Args:
            owner (CalcSheet): Sheet that owns this cell range.
            rng (Any): Range object.
        """
        self.__owner = owner
        if mLo.Lo.is_uno_interfaces(rng, XCellRange):
            self.__range_obj = mCalc.Calc.get_range_obj(cell_range=rng)
            cell_range = rng
        else:
            self.__range_obj = mCalc.Calc.get_range_obj(rng)
            cell_range = mCalc.Calc.get_cell_range(sheet=self.calc_sheet.component, range_obj=self.__range_obj)
        SheetCellRangeComp.__init__(self, cell_range)  # type: ignore
        QiPartial.__init__(self, component=cell_range, lo_inst=mLo.Lo.current_lo)  # type: ignore
        PropPartial.__init__(self, component=cell_range, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    def change_style(self, style_name: str) -> bool:
        """
        Changes style of a range of cells.

        Args:
            style_name (str): Name of style to apply.
            range_obj (RangeObj): Range Object.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        return self.calc_sheet.change_style(style_name=style_name, range_obj=self.__range_obj)

    def clear_cells(self, cell_flags: CellFlagsEnum | None = None) -> bool:
        """
        Clears the specified contents of the cell range

        If ``cell_flags`` is not specified then
        cell range of types ``VALUE``, ``DATETIME`` and ``STRING`` are cleared

        Raises:
            MissingInterfaceError: If XSheetOperation interface cannot be obtained.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_CLEARING` :eventref:`src-docs-cell-event-clearing`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_CLEARED` :eventref:`src-docs-cell-event-cleared`

        Args:
            cell_flags (CellFlagsEnum, optional): Cell flags to clear. Default ``None``.

        Returns:
            bool: True if cells are cleared; Otherwise, False

        Note:
            Events arg for this method have a ``cell`` type of ``XCellRange``.

            Events arg ``event_data`` is a dictionary containing ``cell_flags``.
        """
        if cell_flags is None:
            return mCalc.Calc.clear_cells(sheet=self.calc_sheet.component, range_val=self.__range_obj)
        return mCalc.Calc.clear_cells(
            sheet=self.calc_sheet.component, range_val=self.__range_obj, cell_flags=cell_flags
        )

    def delete_cells(self, is_shift_left: bool) -> bool:
        """
        Deletes cell in a spreadsheet

        Args:
            is_shift_left (bool): If True then cell are shifted left; Otherwise, cells are shifted up.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_DELETING` :eventref:`src-docs-cell-event-deleting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_DELETED` :eventref:`src-docs-cell-event-deleted`

        Returns:
            bool: True if cells are deleted; Otherwise, False

        Note:
            Events args for this method have a ``cell`` type of ``XCellRange``

        Note:
            Event args ``event_data`` is a dictionary containing ``is_shift_left``.
        """
        return mCalc.Calc.delete_cells(
            sheet=self.calc_sheet.component, range_obj=self.__range_obj, is_shift_left=is_shift_left
        )

    def get_cell_series(self) -> XCellSeries:
        """
        Gets the cell series for the current range.

        Returns:
            XCellSeries: Cell series
        """
        return self.qi(XCellSeries, True)

    def get_cell_range(self) -> XCellRange:
        """
        Gets the ``XCellRange`` for the current range.

        Returns:
            XCellRange: Cell series
        """
        return self.qi(XCellRange, True)

    def get_col(self) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        return self.calc_sheet.get_col(range_obj=self.__range_obj)

    def get_row(self) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            range_obj (RangeObj): Range Object

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        return self.calc_sheet.get_row(range_obj=self.__range_obj)

    def get_cell_range_address(self) -> CellRangeAddress:
        """
        Gets the cell range address for the current range.

        Returns:
            CellRangeAddress: Cell range address
        """
        return self.__range_obj.get_cell_range_address()

    def get_range_size(self) -> Size:
        """
        Gets the size of the range.

        Returns:
            ~ooodev.utils.data_type.size.Size: Size, Width is number of Columns and Height is number of Rows
        """
        return mCalc.Calc.get_range_size(range_obj=self.__range_obj)

    def get_range_str(self) -> str:
        """
        Gets the range as a string in format of ``A1:B2`` or ``Sheet1.A1:B2``

        If ``sheet`` is included the format ``Sheet1.A1:B2`` is returned; Otherwise,
        ``A1:B2`` format is returned.

        Returns:
            str: range as string
        """
        return str(self.__range_obj)

    def is_single_cell_range(self) -> bool:
        """
        Gets if a cell address is a single cell or a range

        Returns:
            bool: ``True`` if single cell; Otherwise, ``False``
        """
        return mCalc.Calc.is_single_cell_range(self.get_cell_range_address())

    def is_single_column_range(self) -> bool:
        """
        Gets if a cell address is a single column or multi-column

        Returns:
            bool: ``True`` if single column; Otherwise, ``False``
        """
        return mCalc.Calc.is_single_column_range(self.get_cell_range_address())

    def is_single_row_range(self) -> bool:
        """
        Gets if a cell address is a single row or multi-row

        Returns:
            bool: ``True`` if single row; Otherwise, ``False``
        """
        return mCalc.Calc.is_single_row_range(self.get_cell_range_address())

    def set_array(self, values: Table, styles: Sequence[StyleT] | None = None) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        if styles:
            self.calc_sheet.set_array(values=values, range_obj=self.__range_obj, styles=styles)
        else:
            self.calc_sheet.set_array(values=values, range_obj=self.__range_obj)

    def set_array_range(self, values: Table, styles: Sequence[StyleT] | None = None) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        if styles:
            self.calc_sheet.set_array_range(range_obj=self.__range_obj, values=values, styles=styles)
        else:
            self.calc_sheet.set_array_range(range_obj=self.__range_obj, values=values)

    def set_cell_range_array(self, values: Table, styles: Sequence[StyleT] | None = None) -> None:
        """
        Inserts array of data into spreadsheet

        Args:
            values (Table): A 2-Dimensional array of value such as a list of list or tuple of tuples.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        if styles:
            self.calc_sheet.set_cell_range_array(cell_range=self.component, values=values, styles=styles)
        else:
            self.calc_sheet.set_cell_range_array(cell_range=self.component, values=values)

    def set_style(self, styles: Sequence[StyleT]) -> None:
        """
        Sets style for cell

        Args:
            styles (Sequence[StyleT]): One or more styles to apply to cell.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        mCalc.Calc.set_style_range(sheet=self.calc_sheet.component, range_obj=self.__range_obj, styles=styles)

    def is_merged_cells(self) -> bool:
        """
        Gets is a range of cells is merged.

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            bool: ``True`` if range is merged; Otherwise, ``False``
        """
        return self.calc_sheet.is_merged_cells(range_obj=self.__range_obj)

    def merge_cells(self, center: bool = False) -> None:
        """
        Merges a range of cells

        Args:
            center (bool): Determines if the merge will be a merge and center. Default ``False``.

        Returns:
            None:

        See Also:
            - :py:meth:`.Calc.unmerge_cells`
            - :py:meth:`.Calc.is_merged_cells`
        """
        self.calc_sheet.merge_cells(range_obj=self.__range_obj, center=center)

    def unmerge_cells(self) -> None:
        """
        Removes merging from a range of cells

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            None:

        See Also:
            - :py:meth:`.Calc.merge_cells`
            - :py:meth:`.Calc.is_merged_cells`
        """
        self.calc_sheet.unmerge_cells(range_obj=self.__range_obj)

    def set_val(self, value: Any) -> None:
        """
        Set the value of the very first cell in the range.

        Useful for merged cells.

        Args:
            value (Any): Value to set.
        """
        cell_obj = self.range_obj.start
        cell = mCalcCell.CalcCell(owner=self.calc_sheet, cell=cell_obj)
        cell.set_val(value=value)

    def get_address(self) -> CellRangeAddress:
        """
        Gets Range Address.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        return self.calc_sheet.get_address(range_obj=self.__range_obj)

    def get_array(self) -> TupleArray:
        """
        Gets a 2-Dimensional array of values from a range of cells.

        Returns:
            TupleArray: 2-Dimensional array of values.
        """
        return self.calc_sheet.get_array(range_obj=self.__range_obj)

    def get_float_array(self) -> FloatTable:
        """
        Gets a 2-Dimensional List of floats.

        Returns:
            FloatTable: 2-Dimensional List of floats.
        """
        return self.calc_sheet.get_float_array(range_obj=self.__range_obj)

    def get_val(self) -> Any:
        """
        Get the value of the very first cell in the range.

        Useful for merged cells.

        Returns:
            Any: Value of cell.
        """
        cell_obj = self.range_obj.start
        cell = mCalcCell.CalcCell(owner=self.calc_sheet, cell=cell_obj)
        return cell.get_val()

    def create_cursor(self) -> mCalcCellCursor.CalcCellCursor:
        """
        Creates a cell cursor to travel in the given range context.

        Returns:
            CalcCellCursor: Cell cursor
        """
        return self.calc_sheet.create_cursor_by_range(range_obj=self.__range_obj)

    # region Properties
    @property
    def calc_sheet(self) -> CalcSheet:
        """Sheet that owns this cell."""
        return self.__owner

    @property
    def range_obj(self) -> RangeObj:
        """Range object."""
        return self.__range_obj

    # endregion Properties
