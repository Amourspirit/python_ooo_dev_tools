from __future__ import annotations
from typing import Any, Dict, List, Set, Tuple, cast, overload, Sequence, TYPE_CHECKING
import uno

from com.sun.star.drawing import XDrawPageSupplier
from com.sun.star.sheet import XSheetCellRange
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.table import XCell
from com.sun.star.util import XProtectable

from ooo.dyn.sheet.cell_flags import CellFlagsEnum as CellFlagsEnum

from ooodev.mock import mock_g
from ooodev.adapter.sheet.named_ranges_comp import NamedRangesComp
from ooodev.adapter.sheet.spreadsheet_comp import SpreadsheetComp
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.lo_events import event_ctx
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.lo_events import observe_events
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.office import calc as mCalc
from ooodev.utils import info as mInfo
from ooodev.utils import props as mProps
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.data_type import cell_obj as mCellObj
from ooodev.utils.data_type import range_obj as mRngObj
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.custom_properties_partial import CustomPropertiesPartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.calc import calc_cell_range as mCalcCellRange
from ooodev.calc import calc_cell as mCalcCell
from ooodev.calc import calc_cell_cursor as mCalcCellCursor
from ooodev.calc import calc_table_col as mCalcTableCol
from ooodev.calc import calc_table_row as mCalcTableRow
from ooodev.calc.partial import sheet_cell_partial as mSheetCellPartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial
from ooodev.calc.partial.popup_rng_sel_partial import PopupRngSelPartial

if TYPE_CHECKING:
    from com.sun.star.sheet import SolverConstraint  # struct
    from com.sun.star.sheet import XDataPilotTables
    from com.sun.star.sheet import XGoalSeek
    from com.sun.star.sheet import XScenario
    from com.sun.star.sheet import XSheetCellCursor
    from com.sun.star.table import CellAddress
    from com.sun.star.table import XCellRange
    from com.sun.star.util import XSearchable
    from com.sun.star.util import XSearchDescriptor

    from ooo.dyn.beans.property_value import PropertyValue
    from ooo.dyn.sheet.solver_constraint_operator import SolverConstraintOperator
    from ooo.dyn.table.cell_range_address import CellRangeAddress

    from ooodev.proto.style_obj import StyleT
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.type_var import Row, Column, Table, TupleArray, FloatTable
    from ooodev.calc.calc_doc import CalcDoc
    from ooodev.calc.calc_charts import CalcCharts
    from ooodev.calc.spreadsheet_draw_page import SpreadsheetDrawPage
    from ooodev.calc.cell.sheet_cell_custom_properties import SheetCellCustomProperties


class CalcSheet(
    LoInstPropsPartial,
    SpreadsheetComp,
    mSheetCellPartial.SheetCellPartial,
    EventsPartial,
    QiPartial,
    PropPartial,
    ServicePartial,
    TheDictionaryPartial,
    StylePartial,
    CalcDocPropPartial,
    CalcSheetPropPartial,
    CustomPropertiesPartial,
    PopupRngSelPartial,
):
    """Class for managing Calc Sheet"""

    def __init__(self, owner: CalcDoc, sheet: XSpreadsheet, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (CalcDoc): Owner Document
            sheet (XSpreadsheet): Sheet instance.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        SpreadsheetComp.__init__(self, sheet)  # type: ignore
        EventsPartial.__init__(self)
        QiPartial.__init__(self, component=sheet, lo_inst=self.lo_inst)
        PropPartial.__init__(self, component=sheet, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=sheet, lo_inst=self.lo_inst)
        TheDictionaryPartial.__init__(self)
        StylePartial.__init__(self, component=sheet)
        mSheetCellPartial.SheetCellPartial.__init__(self, owner=self, lo_inst=self.lo_inst)
        CalcDocPropPartial.__init__(self, obj=owner)
        CalcSheetPropPartial.__init__(self, obj=self)
        PopupRngSelPartial.__init__(self, doc=owner)
        self._draw_page = None
        self._charts = None
        self._unique_id = None
        self._named_ranges = None
        self._custom_cell_properties = None
        forms = self.calc_sheet.draw_page.forms.component
        CustomPropertiesPartial.__init__(
            self, forms=forms, form_name="Form_SheetCustomProperties", ctl_name="Sheet_CustomProperties"
        )
        self._init_events()

    def __eq__(self, other: Any) -> bool:
        if not isinstance(other, CalcSheet):
            return False
        return self.calc_sheet.component == other.calc_sheet.component

    # region Events
    def _init_events(self) -> None:
        self._fn_on_range_before_from_obj = self._on_range_before_from_obj
        # self.subscribe_event(GblNamedEvent.RANGE_OBJ_BEFORE_FROM_RANGE, self._fn_on_range_before_from_obj)

    def _on_range_before_from_obj(self, source: Any, args: CancelEventArgs) -> None:
        # set the sheet index when the range object is being created from a range within this class.
        idx = self.get_sheet_index()
        args.event_data["sheet_index"] = idx

    # endregion Events

    # region make_range_obj()
    @overload
    def rng(self, range_name: str) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            range_name (str): Cell range as string.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def rng(self, cell_range: XCellRange) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def rng(self, cr_addr: CellRangeAddress) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            cr_addr (CellRangeAddress): Cell Range Address.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def rng(self, range_obj: mRngObj.RangeObj) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            range_obj (RangeObj): Range Object. If passed in the same RangeObj is returned.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def rng(self, cell_obj: mCellObj.CellObj) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def rng(self, cell_range: XCellRange) -> mRngObj.RangeObj:
        """
        Gets a range Object representing a range.

        Args:
            cell_range (XCellRange): Cell Range.
            sheet (XSpreadsheet): Spreadsheet.

        Returns:
            RangeObj: Range object.
        """
        ...

    @overload
    def rng(self, col_start: int, row_start: int, col_end: int, row_end: int) -> mRngObj.RangeObj:
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

    def rng(self, *args, **kwargs) -> mRngObj.RangeObj:
        """
        Makes a range object.

        Args:
            sheet (XSpreadsheet): Spreadsheet
            range_name (str): Range name such as 'A1:D7'.
            range_obj (RangeObj): Range object.
            start_col (int): Zero-base start column index.
            start_row (int): Zero-base start row index.
            end_col (int): Zero-base end column index.
            end_row (int): Zero-base end row index.

        Returns:
            RangeObj: Range object.
        """
        # use context manage to get this sheet index for the range
        with event_ctx(
            (GblNamedEvent.RANGE_OBJ_BEFORE_FROM_RANGE, self._fn_on_range_before_from_obj),
            source=self,
            lo_observe=True,
        ):
            kargs_len = len(kwargs)
            count = len(args) + kargs_len
            if count == 1:
                val = None
                for v in kwargs.values():
                    val = v
                if val is None and args:
                    val = args[0]

                if val:
                    if mInfo.Info.is_instance(val, mRngObj.RangeObj):
                        return val
                    if mInfo.Info.is_instance(val, str):
                        return mRngObj.RangeObj.from_range(val)
                    if mInfo.Info.is_instance(val, mCellObj.CellObj):
                        return val.get_range_obj()

            range_name = mCalc.Calc.get_range_str(*args, **kwargs)
            return mRngObj.RangeObj.from_range(range_val=range_name)

    # endregion make_range_obj()

    # region get_address()
    @overload
    def get_address(self, *, cell_range: XCellRange) -> CellRangeAddress:
        """
        Gets Range Address.

        Args:
            cell_range (XCellRange): Cell Range.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        ...

    @overload
    def get_address(self, *, range_name: str) -> CellRangeAddress:
        """
        Gets Range Address.

        Args:
            range_name (str): Range name such as 'A1:D7'.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        ...

    @overload
    def get_address(self, *, range_obj: mRngObj.RangeObj) -> CellRangeAddress:
        """
        Gets Range Address.

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        ...

    @overload
    def get_address(self, *, start_col: int, start_row: int, end_col: int, end_row: int) -> CellRangeAddress:
        """
        Gets Range Address.

        Args:
            start_col (int): Zero-base start column index.
            start_row (int): Zero-base start row index.
            end_col (int): Zero-base end column index.
            end_row (int): Zero-base end row index.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        ...

    def get_address(self, **kwargs) -> CellRangeAddress:
        """
        Gets Range Address.

        Args:
            cell_range (XCellRange): Cell Range.
            range_name (str): Range name such as 'A1:D7'.
            range_obj (RangeObj): Range Object.
            start_col (int): Zero-base start column index.
            start_row (int): Zero-base start row index.
            end_col (int): Zero-base end column index.
            end_row (int): Zero-base end row index.

        Returns:
            CellRangeAddress: Cell Range Address.
        """
        sheet_names = {"range_name", "range_obj", "start_col"}
        if kwargs.keys() & sheet_names:
            kwargs["sheet"] = self.component
        return mCalc.Calc.get_address(**kwargs)

    # endregion get_address()

    # region get_cell_address()
    @overload
    def get_cell_address(self, *, cell: XCell) -> CellAddress:
        """
        Gets Cell Address.

        Args:
            cell (XCell): Cell.

        Returns:
            CellAddress: Cell Address.
        """
        ...

    @overload
    def get_cell_address(self, *, cell_name: str) -> CellAddress:
        """
        Gets Cell Address.

        Args:
            cell_name (str): Cell name such as ``A1``.

        Returns:
            CellAddress: Cell Address.
        """
        ...

    @overload
    def get_cell_address(self, *, cell_obj: mCellObj.CellObj) -> CellAddress:
        """
        Gets Cell Address.

        Args:
            cell_obj (CellObj): Cell object.

        Returns:
            CellAddress: Cell Address.
        """
        ...

    @overload
    def get_cell_address(self, *, addr: CellAddress) -> CellAddress:
        """
        Gets Cell Address.

        Args:
            addr (CellAddress): Cell Address.

        Returns:
            CellAddress: Cell Address.
        """
        ...

    @overload
    def get_cell_address(self, *, col: int, row: int) -> CellAddress:
        """
        Gets Cell Address.

        Args:
            col (int): Zero-base column index.
            row (int): Zero-base row index.

        Returns:
            CellAddress: Cell Address.
        """
        ...

    def get_cell_address(self, **kwargs) -> CellAddress:
        """
        Gets Cell Address.

        Args:
            cell (XCell): Cell.
            cell_name (str): Cell name such as ``A1``.
            cell_obj (CellObj): Cell object.
            addr (CellAddress): Cell Address.
            col (int): Zero-base column index.
            row (int): Zero-base row index.

        Returns:
            CellAddress: Cell Address.
        """
        sheet_names = {"cell_name", "cell_obj", "addr", "col"}
        if kwargs.keys() & sheet_names:
            kwargs["sheet"] = self.component
        return mCalc.Calc.get_cell_address(**kwargs)

    # endregion get_cell_address()

    # region get_array()
    @overload
    def get_array(self, *, cell_range: XCellRange) -> TupleArray:
        """
        Gets Array of data from a spreadsheet.

        Args:
            cell_range (XCellRange): Cell range to get data from.

        Returns:
            TupleArray: Resulting data array.
        """
        ...

    @overload
    def get_array(self, *, range_name: str) -> TupleArray:
        """
        Gets Array of data from a spreadsheet.

        Args:
            range_name (str): Range of data to get such as ``A1:E16``.

        Returns:
            TupleArray: Resulting data array.
        """
        ...

    @overload
    def get_array(self, *, range_obj: mRngObj.RangeObj) -> TupleArray:
        """
        Gets Array of data from a spreadsheet.

        Args:
            range_obj (RangeObj): Range object.

        Returns:
            TupleArray: Resulting data array.
        """
        ...

    @overload
    def get_array(self, *, cell_obj: mCellObj.CellObj) -> TupleArray:
        """
        Gets Array of data from a spreadsheet.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            TupleArray: Resulting data array.
        """
        ...

    def get_array(self, **kwargs) -> TupleArray:
        """
        Gets Array of data from a spreadsheet.

        Args:
            cell_range (XCellRange): Cell range to get data from..
            range_name (str): Range of data to get such as ``A1:E16``.
            range_obj (RangeObj): Range object.
            cell_obj (CellObj): Cell Object.

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

    def get_pilot_tables(self) -> XDataPilotTables:
        """
        Gets pivot tables (formerly known as DataPilot) for a sheet.

        Returns:
            XDataPilotTables: Pivot tables
        """
        return mCalc.Calc.get_pilot_tables(self.component)

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
            if "sheet" not in kwargs:
                kwargs["sheet"] = self.component

        # use context manage to get this sheet index for the range
        with event_ctx(
            (GblNamedEvent.RANGE_OBJ_BEFORE_FROM_RANGE, self._fn_on_range_before_from_obj),
            source=self,
            lo_observe=True,
        ):
            range_obj = mCalc.Calc.get_range_obj(**kwargs)
        return mCalcCellRange.CalcCellRange(owner=self, rng=range_obj, lo_inst=self.lo_inst)

    # endregion get_range()

    def get_col_range(self, idx: int) -> mCalcTableCol.CalcTableCol:
        """
        Get Column by index

        Args:
            idx (int): Zero-based column index

        Raises:
            MissingInterfaceError: if unable to find interface

        Returns:
            CalcTableCol: Cell range
        """
        result = mCalc.Calc.get_col_range(self.component, idx)
        return mCalcTableCol.CalcTableCol(owner=self, col_obj=result, lo_inst=self.lo_inst)  # type: ignore

    # region get_row()
    @overload
    def get_row(self, *, calc_cell_range: mCalcCellRange.CalcCellRange) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            calc_cell_range (CalcCellRange): Calc cell range to get row data from.

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        ...

    @overload
    def get_row(self, *, cell_range: XCellRange) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            cell_range (XCellRange): Cell range to get row data from.

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        ...

    @overload
    def get_row(self, *, row_idx: int) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            row_idx (int): Zero base row index such as `0` for row `1`

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        ...

    @overload
    def get_row(self, *, range_name: str) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            range_name (str): Range such as 'A1:A12'

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        ...

    @overload
    def get_row(self, *, cell_obj: mCellObj.CellObj) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            cell_obj (CellObj): Cell Object

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        ...

    @overload
    def get_row(self, *, range_obj: mRngObj.RangeObj) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            range_obj (RangeObj): Range Object

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        ...

    def get_row(self, **kwargs) -> Row:
        """
        Gets a row of data from spreadsheet

        Args:
            calc_cell_range (CalcCellRange): Calc cell range to get row data from.
            cell_range (XCellRange): Cell range to get row data from.
            row_idx (int): Zero base row index such as `0` for row `1`
            range_name (str): Range such as 'A1:A12'
            cell_obj (CellObj): Cell Object
            range_obj (RangeObj): Range Object

        Returns:
            Row: 1-Dimensional List of values on success; Otherwise, None
        """
        if "calc_cell_range" in kwargs:
            arg0 = cast(mCalcCellRange.CalcCellRange, kwargs["calc_cell_range"]).component
            return mCalc.Calc.get_row(arg0)
        elif "cell_range" in kwargs:
            arg0 = cast("XCellRange", kwargs["cell_range"])
            return mCalc.Calc.get_row(arg0)
        kwargs["sheet"] = self.component
        return mCalc.Calc.get_row(**kwargs)

    # endregion get_row()

    def get_row_range(self, idx: int) -> mCalcTableRow.CalcTableRow:
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
        return mCalcTableRow.CalcTableRow(owner=self, row_obj=result, lo_inst=self.lo_inst)  # type: ignore

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
        with LoContext(self.lo_inst):
            addr = mCalc.Calc.get_selected_cell_addr(self.calc_doc.component)
        return self.get_cell(addr=addr)

    def get_selected(self) -> mCalcCellRange.CalcCellRange:
        """
        Gets selected cell range.

        Returns:
            CalcCellRange: Selected cell range
        """
        range_obj = self.get_selected_range()
        return mCalcCellRange.CalcCellRange(owner=self, rng=range_obj, lo_inst=self.lo_inst)

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

        Returns:
            int: Sheet Index
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

    # region get_num()
    @overload
    def get_num(self, *, cell: XCell) -> float:
        """
        Get cell value a float.

        Args:
            cell (XCell): Cell to get value of.

        Returns:
            float: Cell value as float. If cell value cannot be converted then 0.0 is returned.
        """
        ...

    @overload
    def get_num(self, *, cell_name: str) -> float:
        """
        Get cell value a float.

        Args:
            cell_name (str): Cell name such as 'B4'.

        Returns:
            float: Cell value as float. If cell value cannot be converted then 0.0 is returned.
        """
        ...

    @overload
    def get_num(self, *, cell_obj: mCellObj.CellObj) -> float:
        """
        Get cell value a float.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            float: Cell value as float. If cell value cannot be converted then 0.0 is returned.
        """
        ...

    @overload
    def get_num(self, *, addr: CellAddress) -> float:
        """
        Get cell value a float.

        Args:
            addr (CellAddress): Cell Address.

        Returns:
            float: Cell value as float. If cell value cannot be converted then 0.0 is returned.
        """
        ...

    @overload
    def get_num(self, *, col: int, row: int) -> float:
        """
        Get cell value a float.

        Args:
            col (int): Cell zero-base column number.
            row (int): Cell zero-base row number.

        Returns:
            float: Cell value as float. If cell value cannot be converted then 0.0 is returned.
        """
        ...

    def get_num(self, **kwargs) -> float:
        """
        Get cell value a float.

        Args:
            cell (XCell): Cell to get value of.
            cell_name (str): Cell name such as 'B4'.
            cell_obj (CellObj): Cell Object.
            addr (CellAddress): Cell Address.
            col (int): Cell zero-base column number.
            row (int): Cell zero-base row number.

        Returns:
            float: Cell value as float. If cell value cannot be converted then 0.0 is returned.
        """
        sheet_names = {"cell_name", "cell_obj", "addr", "col"}
        if kwargs.keys() & sheet_names:
            kwargs["sheet"] = self.component
        return mCalc.Calc.get_num(**kwargs)

    # endregion get_num()

    # region get_col()
    @overload
    def get_col(self, *, calc_cell_range: mCalcCellRange.CalcCellRange) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        Args:
            calc_cell_range (CalcCellRange): Calc cell range to get column data from.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    @overload
    def get_col(self, *, cell_range: XCellRange) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        Args:
            cell_range (XCellRange): Cell range to get column data from.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    @overload
    def get_col(self, *, col_name: str) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        Args:
            col_name (str): column name such as ``A``.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    @overload
    def get_col(self, *, col_idx: int) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        Args:
            col_idx (int): Zero base column index such as `0` for column ``A``.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    @overload
    def get_col(self, *, range_name: str) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        Args:
            range_name (str): Range such as ``A1:A12``.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    @overload
    def get_col(self, *, cell_obj: mCellObj.CellObj) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    @overload
    def get_col(self, *, range_obj: mRngObj.RangeObj) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        Args:
            range_obj (RangeObj): Range Object.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        ...

    def get_col(self, **kwargs) -> List[Any]:
        """
        Gets a column of data from spreadsheet.

        Args:
            calc_cell_range (CalcCellRange): Calc cell range to get column data from.
            cell_range (XCellRange): Cell range to get column data from.
            col_name (str): column name such as ``A``.
            col_idx (int): Zero base column index such as `0` for column ``A``.
            range_name (str): Range such as ``A1:A12``.
            range_obj (RangeObj): Range Object.
            cell_obj (CellObj): Cell Object.

        Returns:
            List[Any]: 1-Dimensional List.
        """
        if "calc_cell_range" in kwargs:
            arg0 = cast(mCalcCellRange.CalcCellRange, kwargs["calc_cell_range"]).component
            return mCalc.Calc.get_col(arg0)
        elif "cell_range" in kwargs:
            arg0 = cast("XCellRange", kwargs["cell_range"])
            return mCalc.Calc.get_col(arg0)
        kwargs["sheet"] = self.component
        return mCalc.Calc.get_col(**kwargs)

    # endregion get_col()

    # region get_val()
    @overload
    def get_val(self, *, cell: XCell) -> Any:
        """
        Gets cell value.

        Args:
            cell (XCell): cell to get value of.

        Returns:
            Any | None: Cell value cell has a value; Otherwise, ``None``.
        """
        ...

    @overload
    def get_val(self, *, addr: CellAddress) -> Any:
        """
        Gets cell value.

        Args:
            addr (CellAddress): Address of cell.

        Returns:
            Any | None: Cell value cell has a value; Otherwise, ``None``.
        """
        ...

    @overload
    def get_val(self, *, cell_name: str) -> Any:
        """
        Gets cell value.

        Args:
            cell_name (str): Name of cell such as 'B4'.

        Returns:
            Any | None: Cell value cell has a value; Otherwise, ``None``.
        """
        ...

    @overload
    def get_val(self, *, cell_obj: mCellObj.CellObj) -> Any:
        """
        Gets cell value.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            Any | None: Cell value cell has a value; Otherwise, ``None``.
        """
        ...

    @overload
    def get_val(self, *, col: int, row: int) -> Any:
        """
        Gets cell value.

        Args:
            col (int): Cell zero-based column.
            row (int): Cell zero-base row.

        Returns:
            Any: Cell value cell has a value; Otherwise, ``None``.
        """
        ...

    def get_val(self, **kwargs) -> Any | None:
        """
        Gets cell value.

        Args:
            cell (XCell): cell to get value of.
            addr (CellAddress): Address of cell.
            cell_name (str): Name of cell such as 'B4'.
            cell_obj (CellObj): Cell Object.
            col (int): Cell zero-based column.
            row (int): Cell zero-base row.

        Returns:
            Any | None: Cell value cell has a value; Otherwise, ``None``.
        """
        sheet_names = {"addr", "cell_name", "cell_obj", "col"}
        if kwargs.keys() & sheet_names:
            kwargs["sheet"] = self.component
        return mCalc.Calc.get_val(**kwargs)

    # endregion get_val()

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
        return mCalcCellRange.CalcCellRange(owner=self, rng=result, lo_inst=self.lo_inst)

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

    def set_col_width(self, width: int | UnitT, idx: int) -> mCalcTableCol.CalcTableCol | None:
        """
        Sets column width. width is in ``mm``, e.g. ``6``

        Args:
            width (int, UnitT): Width in ``mm`` units or :ref:`proto_unit_obj`.
            idx (int): Index of column.

        Raises:
            CancelEventError: If SHEET_COL_WIDTH_SETTING event is canceled.

        Returns:
            CalcTableCol | None: Column cell range that width is applied on or ``None`` if column width <= 0

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
        return mCalcTableCol.CalcTableCol(self, result)  # type: ignore

    # endregion set_cell_range_array()

    # region change_style()
    @overload
    def change_style(self, style_name: str, cell_range: XCellRange) -> bool:
        """
        Changes style of a range of cells.

        Args:
            style_name (str): Name of style to apply.
            cell_range (XCellRange): Cell range to apply style to.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        ...

    @overload
    def change_style(self, style_name: str, range_name: str) -> bool:
        """
        Changes style of a range of cells.

        Args:
            style_name (str): Name of style to apply.
            range_name (str): Range to apply style to such as ``A1:E23``.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        ...

    @overload
    def change_style(self, style_name: str, range_obj: mRngObj.RangeObj) -> bool:
        """
        Changes style of a range of cells.

        Args:
            style_name (str): Name of style to apply.
            range_obj (RangeObj): Range Object.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        ...

    @overload
    def change_style(self, style_name: str, start_col: int, start_row: int, end_col: int, end_row: int) -> bool:
        """
        Changes style of a range of cells.

        Args:
            style_name (str): Name of style to apply.
            start_col (int): Zero-base start column index.
            start_row (int): Zero-base start row index.
            end_col (int): Zero-base end column index.
            end_row (int): Zero-base end row index.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        ...

    def change_style(self, *args, **kwargs) -> bool:
        """
        Changes style of a range of cells.

        Args:
            style_name (str): Name of style to apply.
            cell_range (XCellRange): Cell range to apply style to.
            range_name (str): Range to apply style to such as ``A1:E23``.
            range_obj (RangeObj): Range Object.
            start_col (int): Zero-base start column index.
            start_row (int): Zero-base start row index.
            end_col (int): Zero-base end column index.
            end_row (int): Zero-base end row index.

        Returns:
            bool: ``True`` if style has been changed; Otherwise, ``False``.
        """
        return mCalc.Calc.change_style(self.component, *args, **kwargs)

    # endregion change_style()

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

    def set_row_height(self, height: int | UnitT, idx: int) -> mCalcTableRow.CalcTableRow | None:
        """
        Sets column width. height is in ``mm``, e.g. 6

        Args:
            height (int, UnitT): Width in ``mm`` units or :ref:`proto_unit_obj`.
            idx (int): Index of Row

        Raises:
            CancelEventError: If SHEET_ROW_HEIGHT_SETTING event is canceled.

        Returns:
            CalcTableRow | None: Row cell range that height is applied on or None if height <= 0

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
        return mCalcTableRow.CalcTableRow(owner=self, row_obj=result, lo_inst=self.lo_inst)  # type: ignore

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
        with LoContext(self.lo_inst):
            # split window dispatches SplitWindow
            # use context manager to observe global events while dispatch is taking place.
            # this will allow CalcDoc to observe the event.
            # eg: doc.subscribe_event("lo_dispatching", on_dispatching)
            with observe_events(self.calc_doc.event_observer):
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
        """
        Go to a cell

        Args:
            cell_name (str): Cell Name such as 'B4'.

        Returns:
            CalcCell: Cell Object.
        """
        ...

    @overload
    def goto_cell(self, cell_obj: mCellObj.CellObj) -> mCalcCell.CalcCell:
        """
        Go to a cell

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            CalcCell: Cell Object.
        """
        ...

    def goto_cell(self, *args, **kwargs) -> mCalcCell.CalcCell:
        """
        Go to a cell

        Args:
            cell_name (str): Cell Name such as 'B4'.
            cell_obj (CellObj): Cell Object.

        Returns:
            CalcCell: Cell Object.

        Attention:
            :py:meth:`~.utils.lo.Lo.dispatch_cmd` method is called along with any of its events.

            Dispatch command is ``GoToCell``.
        """

        def go(obj) -> mCellObj.CellObj:
            frame = self.calc_doc.get_frame()
            s = str(obj)
            props = mProps.Props.make_props(ToPoint=s)
            _ = self.calc_doc.dispatch_cmd(cmd="GoToCell", props=props, frame=frame)
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
    def goal_seek(
        self,
        gs: XGoalSeek,
        cell_name: str | mCellObj.CellObj,
        formula_cell_name: str | mCellObj.CellObj,
        result: int | float,
    ) -> float:
        """
        Calculates a value which gives a specified result in a formula.

        Args:
            gs (XGoalSeek): Goal seeking value for cell
            cell_name (str | CellObj): cell name
            formula_cell_name (str | CellObj): formula cell name
            result (int, float): float or int, result of the goal seek

        Raises:
            GoalDivergenceError: If goal divergence is greater than 0.1

        Returns:
            float: result of the goal seek
        """
        return mCalc.Calc.goal_seek(
            gs=gs, sheet=self.component, cell_name=cell_name, formula_cell_name=formula_cell_name, result=result
        )

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
        self.calc_doc.dispatch_cmd(cmd="Calculate")

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

        Hint:
            - ``CellFlagsEnum`` can be imported from ``ooo.dyn.sheet.cell_flags``
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

        Hint:
            - ``CellFlagsEnum`` can be imported from ``ooo.dyn.sheet.cell_flags``
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

        Hint:
            - ``CellFlagsEnum`` can be imported from ``ooo.dyn.sheet.cell_flags``
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

        Hint:
            - ``CellFlagsEnum`` can be imported from ``ooo.dyn.sheet.cell_flags``
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

        Hint:
            - ``CellFlagsEnum`` can be imported from ``ooo.dyn.sheet.cell_flags``

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

                from ooodev.loader.lo import Lo
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
        return [mCalcCellRange.CalcCellRange(owner=self, rng=x, lo_inst=self.lo_inst) for x in found]

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
        with LoContext(self.lo_inst):
            result = mCalc.Calc.find_function(*args, **kwargs)
        return result

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
        return mCalcCellRange.CalcCellRange(owner=self, rng=found, lo_inst=self.lo_inst)

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

    # region find_used_range()
    @overload
    def find_used_range(self) -> mCalcCellRange.CalcCellRange:
        """
        Find used range

        Returns:
            CalcCellRange: Cell Range.
        """
        ...

    @overload
    def find_used_range(self, range_name: str) -> mCalcCellRange.CalcCellRange:
        """
        Find used range

        Args:
            range_name (str): Range Name such as 'A1:D5'

        Returns:
            CalcCellRange: Cell Range.
        """
        ...

    @overload
    def find_used_range(self, range_obj: mRngObj.RangeObj) -> mCalcCellRange.CalcCellRange:
        """
        Find used range

        Args:
            range_obj (RangeObj): Range Object

        Returns:
            CalcCellRange: Cell Range.
        """
        ...

    @overload
    def find_used_range(self, cr_addr: CellRangeAddress) -> mCalcCellRange.CalcCellRange:
        """
        Find used range

        Args:
            cr_addr (CellRangeAddress): Cell range Address

        Returns:
            CalcCellRange: Cell Range.
        """
        ...

    def find_used_range(self, *args, **kwargs) -> mCalcCellRange.CalcCellRange:
        """
        Find used range

        Args:
            range_name (str): Range Name such as 'A1:D5'
            range_obj (RangeObj): Range Object
            cr_addr (CellRangeAddress): Cell range Address

        Returns:
            CalcCellRange: Cell Range.
        """
        rng = self.find_used_range_obj(*args, **kwargs)
        return mCalcCellRange.CalcCellRange(self, rng)

    # endregion find_used_range()

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
        pro = self.qi(XProtectable, True)
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
        return mCalcCellCursor.CalcCellCursor(owner=self, cursor=cursor, lo_inst=self.lo_inst)

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
        return mCalcCellCursor.CalcCellCursor(owner=self, cursor=cursor, lo_inst=self.lo_inst)

    # endregion create_cursor_by_range()
    @overload
    def set_style_range(self, *, range_name: str, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell range.

        Args:
            range_name (str): Range Name such as ``A1:D5``.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    def set_style_range(self, *, range_obj: mRngObj.RangeObj, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell range.

        Args:
            range_obj (RangeObj): Range Object.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    def set_style_range(self, *, cell_obj: mCellObj.CellObj, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell range.

        Args:
            cell_obj (CellObj): Cell Object.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    def set_style_range(self, *, cr_addr: CellRangeAddress, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell range.

        Args:
            cr_addr (CellRangeAddress): Cell range Address.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    def set_style_range(self, *, cell_range: XCellRange, styles: Sequence[StyleT]) -> None:
        """
        Set style/formatting on cell range.

        Args:
            cell_range (XCellRange): Cell Range. If passed in then the same instance is returned.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    @overload
    def set_style_range(
        self,
        *,
        col_start: int,
        row_start: int,
        col_end: int,
        row_end: int,
        styles: Sequence[StyleT],
    ) -> None:
        """
        Set style/formatting on cell range.

        Args:
            col_start (int): Start Column.
            row_start (int): Start Row.
            col_end (int): End Column.
            row_end (int): End Row.
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:
        """
        ...

    def set_style_range(self, **kwargs) -> None:
        """
        Set style/formatting on cell range.

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
            styles (Sequence[StyleT], optional): One or more styles to apply to cell range.

        Returns:
            None:

        See Also:
            - :ref:`help_calc_format_style_cell`
            - :ref:`help_calc_format_direct_cell`
        """
        if "cell_range" not in kwargs:
            kwargs["sheet"] = self.component
        mCalc.Calc.set_style_range(**kwargs)

    # region make_constraint()
    @overload
    def make_constraint(self, *, num: int | float, op: str, addr: CellAddress) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        Args:
            num (Number): Constraint number such as float or int.
            op (str): Operation such as ``<=``.
            addr (CellAddress): Cell Address.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.
        """
        ...

    @overload
    def make_constraint(
        self, *, num: int | float, op: SolverConstraintOperator, addr: CellAddress
    ) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        Args:
            num (Number): Constraint number such as float or int.
            op (SolverConstraintOperator): Operation such as ``SolverConstraintOperator.EQUAL``.
            addr (CellAddress): Cell Address.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.

        Hint:
            - ``SolverConstraintOperator`` can be imported from ``ooo.dyn.sheet.solver_constraint_operator``
        """

        ...

    @overload
    def make_constraint(self, *, num: int | float, op: str, sheet: XSpreadsheet, cell_name: str) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        Args:
            num (Number): Constraint number such as float or int.
            op (str): Operation such as ``<=``.
            cell_name (str): Cell name such as ``A1``.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.
        """
        ...

    @overload
    def make_constraint(
        self, *, num: int | float, op: str, sheet: XSpreadsheet, cell_obj: mCellObj.CellObj
    ) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        Args:
            num (Number): Constraint number such as float or int.
            op (str): Operation such as ``<=``.
            cell_obj (CellObj): Cell Object.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.
        """
        ...

    @overload
    def make_constraint(
        self, *, num: int | float, op: SolverConstraintOperator, sheet: XSpreadsheet, cell_name: str
    ) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        Args:
            num (Number): Constraint number such as float or int.
            op (SolverConstraintOperator): Operation such as ``SolverConstraintOperator.EQUAL``.
            cell_name (str): Cell name such as ``A1``.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.

        Hint:
            - ``SolverConstraintOperator`` can be imported from ``ooo.dyn.sheet.solver_constraint_operator``
        """
        ...

    def make_constraint(self, **kwargs) -> SolverConstraint:
        """
        Makes a constraint for a solver model.

        Args:
            num (Number): Constraint number such as float or int.
            op (str | SolverConstraintOperator): Operation such as ``<=``.
            addr (CellAddress): Cell Address.
            cell_name (str): Cell name such as ``A1``.
            cell_obj (CellObj): Cell Object.

        Returns:
            SolverConstraint: Solver constraint that can be use in a solver model.

        Hint:
            - ``SolverConstraintOperator`` can be imported from ``ooo.dyn.sheet.solver_constraint_operator``
        """
        sheet_names = {"cell_name", "cell_obj"}
        if kwargs.keys() & sheet_names:
            kwargs["sheet"] = self.component
        return mCalc.Calc.make_constraint(**kwargs)

    # endregion make_constraint()

    # region Properties

    @property
    def name(self) -> str:
        """
        Gets/Sets the sheet Name.

        Returns:
            str: Sheet name
        """
        return self.component.Name  # type: ignore

    @name.setter
    def name(self, value: str) -> None:
        self.component.Name = value  # type: ignore

    sheet_name = name

    @property
    def sheet_index(self) -> int:
        """
        Gets the sheet index.

        Returns:
            int: Sheet index
        """
        return self.component.getRangeAddress().Sheet

    @property
    def draw_page(self) -> SpreadsheetDrawPage[CalcSheet]:
        """
        Gets draw page.

        Returns:
            SpreadsheetDrawPage: Draw Page
        """
        # pylint: disable=import-outside-toplevel
        if self._draw_page is None:
            from ooodev.calc.spreadsheet_draw_page import SpreadsheetDrawPage

            supp = self.qi(XDrawPageSupplier, True)
            draw_page = supp.getDrawPage()
            self._draw_page = SpreadsheetDrawPage(owner=self, component=draw_page, lo_inst=self.lo_inst)
        return self._draw_page  # type: ignore

    @property
    def charts(self) -> CalcCharts:
        """
        Gets charts.

        Returns:
            CalcCharts: Calc Charts
        """
        # pylint: disable=import-outside-toplevel
        if self._charts is None:
            from ooodev.calc.calc_charts import CalcCharts

            self._charts = CalcCharts(owner=self, charts=self.component.getCharts(), lo_inst=self.lo_inst)
        return self._charts

    @property
    def code_name(self) -> str:
        """
        Gets the code name of the sheet.

        If the sheet is renamed this value remains the same. This value is used in macros and is persisted in the document.

        Returns:
            str: Code Name.

        Note:
            The code name may contain any character except for the following: ``[]:*?/\\``.

            Code name are added the the sheet when the sheet is created. If the sheet is renamed the code name will not change.
            Generally the code name will match the original sheet name.

        .. versionadded:: 0.44.1
        """
        return self.component.CodeName  # type: ignore

    @property
    def unique_id(self) -> str:
        """
        Gets the unique name of the sheet.

        Returns:
            str: Unique Name

        Note:
            The unique name is a is different then the sheet name and ``code_name``.
            Unlike the ``code_name`` the unique name will only contain lower case alpha characters.

            The unique name is not added to a sheet until this property is access for the first time.
            After a unique name is added to the sheet it will not change and is persisted with document saves.
            This means that the unique name will not change when the sheet is renamed or moved.

        .. versionadded:: 0.44.1
        """
        if self._unique_id is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.calc.calc_sheet_id import CalcSheetId

            sheet_id = CalcSheetId(sheet=self)
            self._unique_id = sheet_id.id
        return self._unique_id

    @property
    def named_ranges(self) -> NamedRangesComp:
        """
        Named Ranges

        Returns:
            NamedRanges: Named Ranges
        """
        if self._named_ranges is None:
            self._named_ranges = NamedRangesComp(component=self.component.NamedRanges)  # type: ignore
        return self._named_ranges  # type: ignore

    @property
    def custom_cell_properties(self) -> SheetCellCustomProperties:
        """
        Gets the custom property access for the sheet.
        """
        # pylint: disable=import-outside-toplevel
        if self._custom_cell_properties is None:
            from ooodev.calc.cell.sheet_cell_custom_properties import SheetCellCustomProperties

            self._custom_cell_properties = SheetCellCustomProperties(self)
        return self._custom_cell_properties

    # endregion Properties

    # region Static Methods

    @classmethod
    def from_obj(cls, obj: Any, lo_inst: LoInst | None = None) -> CalcSheet | None:
        """
        Creates a CalcSheet from an object.

        Args:
            obj (Any): Object to create CalcSheet from. Can be a CalcSheet, CalcSheetPropPartial, or any object that can be converted to a CalcSheet such as a sheet, cell, or range.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to ``None``.

        Returns:
            CalcSheet: CalcSheet if found; Otherwise, ``None``

        .. versionadded:: 0.46.0
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.calc.calc_doc import CalcDoc

        if mInfo.Info.is_instance(obj, CalcSheetPropPartial):
            return obj.calc_sheet

        calc_doc = CalcDoc.from_obj(obj=obj, lo_inst=lo_inst)
        if calc_doc is None:
            return None

        if hasattr(obj, "component"):
            obj = obj.component

        if not hasattr(obj, "getImplementationName"):
            return None

        sheet = None
        imp_name = obj.getImplementationName()
        if imp_name == "ScTableSheetObj":
            sheet = obj

        if sheet is None:
            if mInfo.Info.is_instance(obj, XSheetCellRange):
                # cell and range object implement XSpreadsheetDocument
                sheet = obj.getSpreadsheet()

        if sheet is not None:
            return cls(owner=calc_doc, sheet=sheet, lo_inst=lo_inst)

        return None

    # endregion Static Methods


if mock_g.FULL_IMPORT:
    from ooodev.calc.spreadsheet_draw_page import SpreadsheetDrawPage
    from ooodev.calc.calc_charts import CalcCharts
    from ooodev.calc.calc_sheet_id import CalcSheetId
    from ooodev.calc.cell.sheet_cell_custom_properties import SheetCellCustomProperties