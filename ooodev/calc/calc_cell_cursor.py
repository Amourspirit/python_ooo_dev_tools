from __future__ import annotations
from typing import Any, cast, List, TYPE_CHECKING, overload
import uno

from com.sun.star.table import XCellRange

from ooodev.adapter.sheet.sheet_cell_cursor_comp import SheetCellCursorComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import calc as mCalc
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial
from ooodev.calc import calc_cell_range as mCalcCellRange
from ooodev.calc import calc_cell as mCalcCell
from ooodev.utils.data_type.cell_obj import CellObj
from ooodev.utils.data_type.range_obj import RangeObj
from ooodev.utils.data_type.range_values import RangeValues
from ooodev.exceptions import ex as mEx

if TYPE_CHECKING:
    from com.sun.star.sheet import SheetCell
    from com.sun.star.table import CellAddress
    from com.sun.star.table import XCell
    from com.sun.star.sheet import XSheetCellCursor
    from ooodev.utils.data_type import cell_obj as mCellObj
    from ooodev.calc.calc_sheet import CalcSheet
else:
    XSheetCellCursor = Any


class CalcCellCursor(
    LoInstPropsPartial,
    SheetCellCursorComp,
    QiPartial,
    PropPartial,
    StylePartial,
    ServicePartial,
    CalcDocPropPartial,
    CalcSheetPropPartial,
    TheDictionaryPartial,
):
    def __init__(self, owner: CalcSheet, cursor: XSheetCellCursor, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        SheetCellCursorComp.__init__(self, cursor)  # type: ignore
        QiPartial.__init__(self, component=cursor, lo_inst=self.lo_inst)  # type: ignore
        PropPartial.__init__(self, component=cursor, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=cursor)
        ServicePartial.__init__(self, component=cursor, lo_inst=self.lo_inst)
        CalcSheetPropPartial.__init__(self, obj=owner.calc_sheet)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)
        TheDictionaryPartial.__init__(self)

    def find_used_cursor(self) -> mCalcCellRange.CalcCellRange:
        """
        Find used cursor

        Raises:
            MissingInterfaceError: if unable to find interface

        Returns:
            CalcCellRange: Cell range
        """
        found = mCalc.Calc.find_used_cursor(self.component)
        return mCalcCellRange.CalcCellRange(owner=self.calc_sheet, rng=found, lo_inst=self.lo_inst)

    def find_used_range_obj(self, content_flags: int = 23) -> RangeObj:
        """
        Finds used range object.

        The used range is found by querying the current range for content specified by the ``content_flags``.

        Args:
            content_flags (int, optional): CellFlags. Defaults to 23.

        Raises:
            CellRangeError: If unable to get used range object

        Returns:
            RangeObj: The Range object that represents the used range.

        Note:
            Default ``CellFlags`` is: ``CellFlags.FORMULA | CellFlags.VALUE | CellFlags.DATETIME | CellFlags.STRING``

            ``CellFlags`` can be imported from ``com.sun.star.sheet``.

        See Also:
            `API CellFlags <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1sheet_1_1CellFlags.html>`_

        .. versionadded:: 0.47.7
        """
        try:
            # content_flags = CellFlags.FORMULA | CellFlags.VALUE | CellFlags.DATETIME | CellFlags.STRING
            cell_range = self.qi(XCellRange, True)
            rng_obj = self.calc_doc.range_converter.get_range_obj(cell_range=cell_range)

            cursor = self.calc_sheet.create_cursor_by_range(range_obj=rng_obj)
            q_result = cursor.component.queryContentCells(content_flags)
            if q_result is None:
                raise mEx.CellRangeError("Error getting used range object: queryContentCells() returned None")
            cells = q_result.getCells()
            if cells is None:
                raise mEx.CellRangeError("Error getting used range object: getCells() returned None")
            if not cells.hasElements():
                raise mEx.CellRangeError("Error getting used range object: getCells() has no elements")
            enum = cells.createEnumeration()
            if enum is None:
                raise mEx.CellRangeError("Error getting used range object: createEnumeration() returned None")
            sheet_cells: List[CellObj] = []
            while enum.hasMoreElements():
                sc = cast("SheetCell", enum.nextElement())
                sheet_cells.append(CellObj.from_cell(sc.getCellAddress()))

            if len(sheet_cells) < 2:
                raise mEx.CellRangeError(
                    f"Error getting used range object: Not enough cells found. Minimum is 2. Found: {len(sheet_cells)}"
                )
            sheet_cells.sort()
            cell_start = sheet_cells[0]
            cell_end = sheet_cells[-1]
            addr_start = cell_start.get_cell_values()
            addr_end = cell_end.get_cell_values()
            result = RangeValues(
                col_start=addr_start.col,
                row_start=addr_start.row,
                col_end=addr_end.col,
                row_end=addr_end.row,
                sheet_idx=rng_obj.sheet_idx,
            )
            # cursor.component.gotoStartOfUsedArea(False) and gotoEndOfUsedArea(True) are not working
            # correctly. The goto methods go outside the bounds of the range.
            return RangeObj.from_range(result)
        except mEx.CellRangeError:
            raise
        except Exception as e:
            raise mEx.CellRangeError(f"Error getting used range object: {e}") from e

    def get_calc_cell_range(self) -> mCalcCellRange.CalcCellRange:
        """
        Get calc cell range

        Returns:
            CalcCellRange: Cell range
        """
        cell_range = self.qi(XCellRange, True)
        return mCalcCellRange.CalcCellRange(owner=self.calc_sheet, rng=cell_range, lo_inst=self.lo_inst)

    # region get_cell_by_position()
    @overload
    def get_cell_by_position(self) -> mCalcCell.CalcCell:
        """
        Get current active cell

        Returns:
            CalcCell: Cell
        """
        ...

    @overload
    def get_cell_by_position(self, cell_name: str) -> mCalcCell.CalcCell:
        """
        Get cell by position

        Args:
            cell_name (str): Cell name.

        Returns:
            CalcCell: Cell
        """
        ...

    @overload
    def get_cell_by_position(self, addr: CellAddress) -> mCalcCell.CalcCell:
        """
        Get cell by position

        Args:
            addr (CellAddress): Cell Address.

        Returns:
            CalcCell: Cell
        """
        ...

    @overload
    def get_cell_by_position(self, cell: XCell) -> mCalcCell.CalcCell:
        """
        Get cell by position

        Args:
            cell (XCell): Cell.

        Returns:
            CalcCell: Cell
        """
        ...

    @overload
    def get_cell_by_position(self, cell_obj: mCellObj.CellObj) -> mCalcCell.CalcCell:
        """
        Get cell by position

        Args:
            cell_obj (CellObj): Cell Object. If passed in the same CellObj is returned.

        Returns:
            CalcCell: Cell
        """
        ...

    @overload
    def get_cell_by_position(self, col: int, row: int) -> mCalcCell.CalcCell:
        """
        Get cell by position

        Args:
            col (int): Zero-based column index.
            row (int): Zero-based row index.

        Returns:
            CalcCell: Cell
        """
        ...

    def get_cell_by_position(self, *args, **kwargs) -> mCalcCell.CalcCell:
        """
        Get cell by position

        Args:
            cell_name (str): Cell name.
            addr (CellAddress): Cell Address.
            cell (XCell): Cell.
            cell_obj (CellObj): Cell Object. If passed in the same CellObj is returned.
            col (int): Zero-based column index.
            row (int): Zero-based row index.

        Returns:
            CalcCell: Cell
        """
        with LoContext(self.lo_inst):
            cell_obj = mCalc.Calc.get_cell_obj(*args, **kwargs)
        # x_cell = self.component.getCellByPosition(cell_obj.col_obj.index, cell_obj.row_obj.index)
        return mCalcCell.CalcCell(owner=self.calc_sheet, cell=cell_obj, lo_inst=self.lo_inst)

    # endregion get_cell_by_position()

    # region Cursor Move Movement
    def go_to_start(self) -> CalcCellCursor:
        """
        Go to start.

        Points the cursor to a single cell which is the beginning of a contiguous series of (filled) cells.

        Returns:
            CalcCellCursor: New instance of CalcCellCursor
        """
        self.component.gotoStart()
        with LoContext(self.lo_inst):
            cell = mCalc.Calc.get_cell_obj()
        return self.calc_sheet.create_cursor_by_range(cell_obj=cell)

    def go_to_next(self) -> CalcCellCursor:
        """
        Go to next.

        Points the cursor to the next unprotected cell.
        If the sheet is not protected, this is the next cell to the right.

        Returns:
            CalcCellCursor: New instance of CalcCellCursor
        """
        self.component.gotoNext()
        with LoContext(self.lo_inst):
            cell = mCalc.Calc.get_cell_obj()
        return self.calc_sheet.create_cursor_by_range(cell_obj=cell)

    def go_to_previous(self) -> CalcCellCursor:
        """
        Go to next.

        Points the cursor to the previous unprotected cell.
        If the sheet is not protected, this is the next cell to the left.

        Returns:
            CalcCellCursor: New instance of CalcCellCursor
        """
        self.component.gotoPrevious()
        with LoContext(self.lo_inst):
            cell = mCalc.Calc.get_cell_obj()
        return self.calc_sheet.create_cursor_by_range(cell_obj=cell)

    def go_to_end(self) -> CalcCellCursor:
        """
        Go to end.

        Points the cursor to a single cell which is the end of a contiguous series of (filled) cells.

        Returns:
            CalcCellCursor: New instance of CalcCellCursor
        """
        self.component.gotoEnd()
        with LoContext(self.lo_inst):
            cell = mCalc.Calc.get_cell_obj()
        return self.calc_sheet.create_cursor_by_range(cell_obj=cell)

    def go_to_offset(self, col_offset: int, row_offset: int) -> CalcCellCursor:
        """
        Go to next.

        Moves the origin of the cursor relative to the current position.

        Returns:
            CalcCellCursor:  New instance of CalcCellCursor
        """
        self.component.gotoOffset(col_offset, row_offset)
        with LoContext(self.lo_inst):
            cell = mCalc.Calc.get_cell_obj()
        return self.calc_sheet.create_cursor_by_range(cell_obj=cell)

    def go_to_start_of_used_area(self, expand: bool = False) -> CalcCellCursor:
        """
        Go to end.

        points the cursor to the start of the used area.

        Args:
            expand (bool): If ``True`` then the used area is expanded to the left and up if necessary.
                ``False``sets size of the cursor to a single cell. Default is ``False``.

        Returns:
            CalcCellCursor: New instance of CalcCellCursor
        """
        self.component.gotoStartOfUsedArea(expand)
        if expand:
            with LoContext(self.lo_inst):
                range_obj = mCalc.Calc.find_used_range_obj(sheet=self.calc_sheet.component)
            return self.calc_sheet.create_cursor_by_range(range_obj=range_obj)
        with LoContext(self.lo_inst):
            cell = mCalc.Calc.get_cell_obj()
        return self.calc_sheet.create_cursor_by_range(cell_obj=cell)

    def go_to_end_of_used_area(self, expand: bool = False) -> CalcCellCursor:
        """
        Go to end.

        points the cursor to the end of the used area.

        Args:
            expand (bool): If ``True`` then the used area is expanded to the right and down if necessary.
                ``False``sets size of the cursor to a single cell. Default is ``False``.

        Returns:
            CalcCellCursor: New instance of CalcCellCursor
        """
        self.component.gotoEndOfUsedArea(expand)
        if expand:
            with LoContext(self.lo_inst):
                range_obj = mCalc.Calc.find_used_range_obj(sheet=self.calc_sheet.component)
            return self.calc_sheet.create_cursor_by_range(range_obj=range_obj)
        with LoContext(self.lo_inst):
            cell = mCalc.Calc.get_cell_obj()
        return self.calc_sheet.create_cursor_by_range(cell_obj=cell)

    # endregion Cursor Move Movement
