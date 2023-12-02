from __future__ import annotations
from typing import Any, Sequence, TYPE_CHECKING
import uno

if TYPE_CHECKING:
    from ooo.dyn.table.cell_range_address import CellRangeAddress
    from .calc_sheet import CalcSheet
else:
    CellRangeAddress = object

from ooodev.proto.style_obj import StyleT
from ooodev.utils.type_var import Table
from ooodev.office import calc as mCalc
from ooodev.adapter.sheet.sheet_cell_range_comp import SheetCellRangeComp


class CalcCellRange(SheetCellRangeComp):
    def __init__(self, owner: CalcSheet, range: Any) -> None:
        self.__owner = owner
        self.__range_obj = mCalc.Calc.get_range_obj(range)
        cell_range = mCalc.Calc.get_cell_range(sheet=self.calc_sheet.component, range_obj=self.__range_obj)
        super().__init__(cell_range)  # type: ignore

    def clear_cells(self) -> bool:
        """
        Clears the specified contents of the cell range

        If cell_flags is not specified then
        cell range of types ``VALUE``, ``DATETIME`` and ``STRING`` are cleared

        Raises:
            MissingInterfaceError: If XSheetOperation interface cannot be obtained.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_CLEARING` :eventref:`src-docs-cell-event-clearing`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.CELLS_CLEARED` :eventref:`src-docs-cell-event-cleared`

        Returns:
            bool: True if cells are cleared; Otherwise, False

        Note:
            Events arg for this method have a ``cell`` type of ``XCellRange``.

            Events arg ``event_data`` is a dictionary containing ``cell_flags``.
        """
        return mCalc.Calc.clear_cells(sheet=self.calc_sheet.component, range_val=self.__range_obj)

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

    def get_cell_range_address(self) -> CellRangeAddress:
        """
        Gets the cell range address for the current range.

        Returns:
            CellRangeAddress: Cell range address
        """
        return self.__range_obj.get_cell_range_address()

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

    # region Properties
    @property
    def calc_sheet(self) -> CalcSheet:
        """Sheet that owns this cell."""
        return self.__owner

    # endregion Properties
