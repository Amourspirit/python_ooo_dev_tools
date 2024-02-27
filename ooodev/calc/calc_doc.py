from __future__ import annotations
from typing import Any, cast, List, Tuple, overload, Sequence, TYPE_CHECKING

# pylint: wrong-import-position
import uno

from com.sun.star.drawing import XDrawPagesSupplier
from com.sun.star.frame import XModel
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheets
from com.sun.star.sheet import XSpreadsheetDocument

from ooodev.adapter.sheet.spreadsheet_document_comp import SpreadsheetDocumentComp
from ooodev.dialog.partial.create_dialog_partial import CreateDialogPartial
from ooodev.events.args.calc.sheet_args import SheetArgs
from ooodev.events.args.calc.sheet_cancel_args import SheetCancelArgs
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.calc_named_event import CalcNamedEvent
from ooodev.events.event_singleton import _Events
from ooodev.events.lo_events import observe_events
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import calc as mCalc
from ooodev.utils import gen_util as mGenUtil
from ooodev.utils import gui as mGUI
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.utils import view_state as mViewState
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.data_type import range_obj as mRngObj
from ooodev.loader.inst.doc_type import DocType
from ooodev.loader.inst.service import Service as LoService
from ooodev.utils.kind.zoom_kind import ZoomKind
from ooodev.utils.partial.dispatch_partial import DispatchPartial
from ooodev.utils.partial.doc_io_partial import DocIoPartial
from ooodev.utils.partial.gui_partial import GuiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.calc import calc_sheet as mCalcSheet
from ooodev.calc import calc_sheets as mCalcSheets
from ooodev.calc import calc_sheet_view as mCalcSheetView
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.spreadsheet_draw_pages import SpreadsheetDrawPages

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.frame import XController
    from com.sun.star.sheet import XViewPane
    from com.sun.star.style import XStyle
    from com.sun.star.table import CellAddress
    from com.sun.star.table import XCellRange
    from ooo.dyn.sheet.general_function import GeneralFunction
    from ooo.dyn.table.cell_range_address import CellRangeAddress
    from ooodev.loader.inst.lo_inst import LoInst
else:
    CellRangeAddress = Any
    SpreadsheetDocument = Any


class CalcDoc(
    LoInstPropsPartial,
    SpreadsheetDocumentComp,
    QiPartial,
    PropPartial,
    GuiPartial,
    ServicePartial,
    StylePartial,
    EventsPartial,
    DocIoPartial["CalcDoc"],
    CreateDialogPartial,
    DispatchPartial,
    CalcDocPropPartial,
):
    """Defines a Calc Document"""

    DOC_TYPE: DocType = DocType.CALC

    def __init__(self, doc: XSpreadsheetDocument, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            doc (XSpreadsheetDocument): UNO object the supports ``com.sun.star.sheet.SpreadsheetDocument`` service.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.

        Raises:
            NotSupportedDocumentError: If not a valid Calc document.

        Returns:
            None:
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not mInfo.Info.is_doc_type(doc, LoService.CALC):
            raise mEx.NotSupportedDocumentError("Document is not a Calc document")
        SpreadsheetDocumentComp.__init__(self, doc)  # type: ignore
        QiPartial.__init__(self, component=doc, lo_inst=self.lo_inst)
        PropPartial.__init__(self, component=doc, lo_inst=self.lo_inst)
        GuiPartial.__init__(self, component=doc, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=doc, lo_inst=self.lo_inst)
        EventsPartial.__init__(self)
        StylePartial.__init__(self, component=doc)
        DocIoPartial.__init__(self, owner=self, lo_inst=self.lo_inst)
        CreateDialogPartial.__init__(self, lo_inst=self.lo_inst)
        DispatchPartial.__init__(self, lo_inst=self.lo_inst, events=self)
        CalcDocPropPartial.__init__(self, obj=self)
        self._sheets = None
        self._draw_pages = None
        self._current_controller = None

    # region context manage
    def __enter__(self) -> CalcDoc:
        self.lock_controllers()
        return self

    def __exit__(self, *exc) -> None:
        self.unlock_controllers()

    # endregion context manage

    def create_cell_style(self, style_name: str) -> XStyle:
        """
        Creates a style

        Args:
            style_name (str): Style name

        Raises:
            Exception: if unable to create style.

        Returns:
            XStyle: Newly created style
        """
        with LoContext(self.lo_inst):
            result = mCalc.Calc.create_cell_style(doc=self.component, style_name=style_name)
        return result

    def call_fun(self, func_name: str, *args: Any) -> Any:
        """
        Execute a Calc function by its (English) name and based on the given arguments

        Args:
            func_name (str): the English name of the function to execute
            args: (Any): the arguments of the called function.
                Each argument must be either a string, a numeric value
                or a sequence of sequences ( tuples or list ) combining those types.

        Returns:
            Any: The (string or numeric) value or the array of arrays returned by the call to the function
                When the arguments contain arrays, the function is executed as an array function
                Wrong arguments generate an error
        """
        with LoContext(self.lo_inst):
            result = mCalc.Calc.call_fun(func_name, *args)
        return result

    def compute_function(self, fn: GeneralFunction | str, cell_range: XCellRange) -> float:
        """
        Computes a Calc Function

        Args:
            fn (GeneralFunction | str): Function to calculate, GeneralFunction Enum value or String such as 'SUM' or 'MAX'
            cell_range (XCellRange): Cell range to apply function on.

        Returns:
            float: result of function if successful. If there is an error then 0.0 is returned.

        See Also:
            :ref:`ch23_gen_func`

        Hint:
            - ``GeneralFunction`` can be imported from ``ooo.dyn.sheet.general_function``
        """
        return mCalc.Calc.compute_function(fn=fn, cell_range=cell_range)

    def get_function_names(self) -> List[str] | None:
        """
        Get a list of all function names.

        Returns:
            List[str] | None: List of function names if found; Otherwise, ``None``.
        """
        with LoContext(self.lo_inst):
            result = mCalc.Calc.get_function_names()
        return result

    # region find_function()
    @overload
    def find_function(self, func_nm: str) -> Tuple[PropertyValue] | None:
        """
        Finds a function.

        Args:
            func_nm (str): function name.

        Returns:
            Tuple[PropertyValue] | None: Function properties as tuple on success; Otherwise, ``None``.
        """
        ...

    @overload
    def find_function(self, idx: int) -> Tuple[PropertyValue] | None:
        """
        Finds a function.

        Args:
            idx (int): Index of function.

        Returns:
            Tuple[PropertyValue] | None: Function properties as tuple on success; Otherwise, ``None``.
        """
        ...

    def find_function(self, *args, **kwargs) -> Tuple[PropertyValue, ...] | None:
        """
        Finds a function.

        Args:
            func_nm (str): function name.
            idx (int): Index of function.

        Returns:
            Tuple[PropertyValue, ...] | None: Function properties as tuple on success; Otherwise, ``None``.
        """
        with LoContext(self.lo_inst):
            result = mCalc.Calc.find_function(*args, **kwargs)
        return result

    # endregion find_function()

    # region close_doc()
    @overload
    def close_doc(self) -> None:
        """
        Closes document.

        Returns:
            None:
        """
        ...

    @overload
    def close_doc(self, deliver_ownership: bool) -> None:
        """
        Closes document.

        Args:
            deliver_ownership (bool): If ``True`` delegates the ownership of this closing object to
                anyone which throw the CloseVetoException. Default is ``False``.

        Returns:
            None:
        """
        ...

    def close_doc(self, deliver_ownership=False) -> None:
        """
        Closes document.

        Args:
            deliver_ownership (bool, optional): If ``True`` ownership is delivered to caller. Default ``True``.
                ``True`` delegates the ownership of this closing object to anyone which throw the CloseVetoException.
                This new owner has to close the closing object again if his still running processes will be finished.
                ``False`` let the ownership at the original one which called the close() method.
                They must react for possible CloseVetoExceptions such as when document needs saving and
                try it again at a later time. This can be useful for a generic UI handling.

        Returns:
            None:

        Note:
            If ``deliver_ownership`` is ``True`` then new owner has to close the closing object again if his still running
            processes will be finished.

            ``False`` let the ownership at the original one which called the close() method.
            They must react for possible CloseVetoExceptions such as when document needs saving
            and try it again at a later time. This can be useful for a generic UI handling.

        Attention:
            :py:meth:`~.utils.lo.Lo.close` method is called along with any of its events.
        """

        self.close(deliver_ownership=deliver_ownership)

    # endregion close_doc()

    def _on_io_saving(self, event_args: CancelEventArgs) -> None:
        """
        Event called before document is saved.

        Args:
            event_args (CancelEventArgs): Event data.

        Raises:
            CancelEventError: If event is canceled.
        """
        event_args.event_data["doc"] = self.component
        self.trigger_event(CalcNamedEvent.DOC_SAVING, event_args)

    def _on_io_saved(self, event_args: EventArgs) -> None:
        """
        Event called after document is saved.

        Args:
            event_args (EventArgs): Event data.
        """
        self.trigger_event(CalcNamedEvent.DOC_SAVED, event_args)

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
        with LoContext(self.lo_inst):
            result = mCalc.Calc.get_selected_addr(doc=self.component)
        return result

    def get_selected_cell_addr(self) -> CellAddress:
        """
        Gets the cell address of current selected cell of the active sheet.

        Raises:
            CellError: if active selection is not a single cell

        Returns:
            CellAddress: Cell Address

        Note:
            CellAddress returns Zero-base values.
            For instance: Cell ``B4`` has Column value of ``1`` and Row value of ``3``

        See Also:
            - :py:meth:`~.Calc.get_selected_cell`
            - :py:meth:`~.Calc.get_selected_addr`
            - :py:meth:`~.Calc.set_selected_addr`
        """
        with LoContext(self.lo_inst):
            result = mCalc.Calc.get_selected_cell_addr(self.component)
        return result

    def get_selected_range(self) -> mRngObj.RangeObj:
        """
        Gets select cell range

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            model (XModel): model used to access sheet

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
        return mCalc.Calc.get_selected_range(doc=self.component)

    # region    get_sheet()
    @overload
    def get_sheet(self) -> mCalcSheet.CalcSheet:
        """
        Gets a sheet of spreadsheet document

        Returns:
            CalcSheet: Spreadsheet at index.
        """
        ...

    @overload
    def get_sheet(self, idx: int) -> mCalcSheet.CalcSheet:
        """
        Gets a sheet of spreadsheet document

        Args:
            idx (int): The Zero-based index of the sheet. Idx can be a negative value to index from the end of the list.
                For example, -1 will return the last sheet.

        Returns:
            CalcSheet: Spreadsheet at index.
        """
        ...

    @overload
    def get_sheet(self, sheet_name: str) -> mCalcSheet.CalcSheet:
        """
        Gets a sheet of spreadsheet document

        Args:
            sheet_name (str, optional): Name of spreadsheet

        Returns:
            CalcSheet: Spreadsheet at index.
        """
        ...

    def get_sheet(self, *args, **kwargs) -> mCalcSheet.CalcSheet:
        """
        Gets a sheet of spreadsheet document

        Args:
            idx (int, optional): Zero based index of spreadsheet. Idx can be a negative value to index from the end of the list.
                For example, -1 will return the last element.
            sheet_name (str, optional): Name of spreadsheet

        Raises:
            MissingNameError: If spreadsheet is not found by name.
            IndexError: If spreadsheet is not found by index.

        Returns:
            CalcSheet: Spreadsheet at index.
        """
        return self.sheets.get_sheet(*args, **kwargs)

    # endregion get_sheet()

    # region insert_sheet
    @overload
    def insert_sheet(self, name: str) -> mCalcSheet.CalcSheet:
        """
        Inserts a spreadsheet into the end of the documents sheet collection.

        Args:
            name (str): Name of sheet to insert
        """
        ...

    @overload
    def insert_sheet(self, name: str, idx: int) -> mCalcSheet.CalcSheet:
        """
        Inserts a spreadsheet into document sheet collections.

        Args:
            name (str): Name of sheet to insert
            idx (int): zero-based index position of the sheet to insert.
                Can be a negative value to insert from the end of the collection.
                Default is ``-1`` which inserts at the end of the collection.
        """
        ...

    def insert_sheet(self, name: str, idx: int = -1) -> mCalcSheet.CalcSheet:
        """
        Inserts a spreadsheet into document sheet collections.

        Args:
            name (str): Name of sheet to insert
            idx (int, optional): zero-based index position of the sheet to insert.
                Can be a negative value to insert from the end of the collection.
                Default is ``-1`` which inserts at the end of the collection.

        Raises:
            Exception: If unable to insert spreadsheet
            CancelEventError: If SHEET_INSERTING event is canceled

        Returns:
            CalcSheet: The newly inserted sheet

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_INSERTING` :eventref:`src-docs-sheet-event-inserting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_INSERTED` :eventref:`src-docs-sheet-event-inserted`
        """
        index = self._get_index(idx=idx, allow_greater=True)
        sheet = mCalc.Calc.insert_sheet(doc=self.component, name=name, idx=index)
        return mCalcSheet.CalcSheet(owner=self, sheet=sheet, lo_inst=self.lo_inst)

    # endregion insert_sheet

    def _get_index(self, idx: int, allow_greater: bool = False) -> int:
        """
        Gets the index.

        Args:
            idx (int): Index of sheet. Can be a negative value to index from the end of the list.
            allow_greater (bool, optional): If True and index is greater then the number of
                sheets then the index becomes the next index if sheet were appended. Defaults to False.

        Returns:
            int: Index value.
        """
        count = len(self.sheets)
        return mGenUtil.Util.get_index(idx, count, allow_greater)

    def freeze(self, num_cols: int, num_rows: int) -> None:
        """
        Freezes spreadsheet columns and rows

        Args:
            num_cols (int): Number of columns to freeze
            num_rows (int): Number of rows to freeze

        Returns:
            None:

        See Also:

            - :ref:`ch23_freezing_rows`
            - :py:meth:`~.Calc.freeze_rows`
            - :py:meth:`~.Calc.freeze_cols`
            - :py:meth:`~.Calc.unfreeze`
        """
        with LoContext(self.lo_inst):
            mCalc.Calc.freeze(doc=self.component, num_cols=num_cols, num_rows=num_rows)

    def freeze_cols(self, num_cols: int) -> None:
        """
        Freezes spreadsheet columns

        Args:
            num_cols (int): Number of columns to freeze

        Returns:
            None:

        See Also:

            - :ref:`ch23_freezing_rows`
            - :py:meth:`~.Calc.freeze`
            - :py:meth:`~.Calc.freeze_rows`
            - :py:meth:`~.Calc.unfreeze`
        """
        with LoContext(self.lo_inst):
            mCalc.Calc.freeze(doc=self.component, num_cols=num_cols, num_rows=0)

    def freeze_rows(self, num_rows: int) -> None:
        """
        Freezes spreadsheet rows

        Args:
            num_rows (int): Number of rows to freeze

        Returns:
            None:

        See Also:

            - :ref:`ch23_freezing_rows`
            - :py:meth:`~.Calc.freeze`
            - :py:meth:`~.Calc.freeze_cols`
            - :py:meth:`~.Calc.unfreeze`
        """
        with LoContext(self.lo_inst):
            mCalc.Calc.freeze(doc=self.component, num_cols=0, num_rows=num_rows)

    # region    remove_sheet()

    @overload
    def remove_sheet(self, sheet_name: str) -> bool:
        """
        Removes a sheet from document

        Args:
            sheet_name (str): Name of sheet to remove

        Returns:
            bool: True of sheet was removed; Otherwise, False
        """
        ...

    @overload
    def remove_sheet(self, idx: int) -> bool:
        """
        Removes a sheet from document

        Args:
            idx (int): Zero based index of sheet to remove.

        Returns:
            bool: True of sheet was removed; Otherwise, False
        """
        ...

    def remove_sheet(self, *args, **kwargs) -> bool:
        """
        Removes a sheet from document

        Args:
            sheet_name (str): Name of sheet to remove
            idx (int): Zero based index of sheet to remove.
                Can be a negative value to insert from the end of the list.

        Returns:
            bool: True of sheet was removed; Otherwise, False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_REMOVING` :eventref:`src-docs-sheet-event-removing`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_REMOVED` :eventref:`src-docs-sheet-event-removed`

        Note:
            Event args ``event_data`` is set to a dictionary.
            If ``idx`` is available then args ``event_data["fn_type"]`` is set to a value ``idx``; Otherwise, set to a value ``name``.
        """

        kargs_len = len(kwargs)
        count = len(args) + kargs_len
        if count == 0:
            raise TypeError("remove_sheet() missing 1 required positional argument: 'sheet_name' or 'idx'")
        if "sheet_name" in kwargs:
            return mCalc.Calc.remove_sheet(doc=self.component, sheet_name=kwargs["sheet_name"])
        if "idx" in kwargs:
            idx = self._get_index(kwargs["idx"])
            return mCalc.Calc.remove_sheet(doc=self.component, idx=idx)
        if kwargs:
            raise TypeError("remove_sheet() got an unexpected keyword argument")
        if count != 1:
            raise TypeError("remove_sheet() got an invalid number of arguments")
        arg = args[0]
        if isinstance(arg, int):
            idx = self._get_index(arg)
            return mCalc.Calc.remove_sheet(doc=self.component, idx=idx)
        if isinstance(arg, str):
            return mCalc.Calc.remove_sheet(doc=self.component, sheet_name=arg)
        raise TypeError("get_sheet() got an invalid argument")

    # endregion remove_sheet()

    def move_sheet(self, name: str, idx: int) -> bool:
        """
        Moves a sheet in a spreadsheet document

        Args:
            name (str): Name of sheet to move
            idx (int): The zero based index to move sheet into.

        Returns:
            bool: True on success; Otherwise, False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_MOVING` :eventref:`src-docs-sheet-event-moving`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_MOVED` :eventref:`src-docs-sheet-event-moved`
        """
        index = self._get_index(idx=idx)
        cargs = SheetCancelArgs(self.move_sheet.__qualname__)
        cargs.doc = self.component
        cargs.name = name
        cargs.index = index
        _Events().trigger(CalcNamedEvent.SHEET_MOVING, cargs)
        if cargs.cancel:
            return False
        name = cargs.name
        index = cargs.index
        num_sheets = len(self.sheets)
        result = False
        if index < 0 or index >= num_sheets:
            mLo.Lo.print(f"Index {index} is out of range.")
        else:
            self.sheets.component.moveByName(name, index)
            result = True
        if result:
            _Events().trigger(CalcNamedEvent.SHEET_MOVED, SheetArgs.from_args(cargs))
        return result

    def get_sheet_names(self) -> Tuple[str, ...]:
        """
        Gets names of all existing spreadsheets in the spreadsheet document.

        Returns:
            Tuple[str, ...]: Tuple of sheet names.
        """
        sheets = self.component.getSheets()
        return sheets.getElementNames()

    def get_sheets(self) -> XSpreadsheets:
        """
        Gets all existing spreadsheets in the spreadsheet document.

        Returns:
            XSpreadsheets: document sheets
        """
        return self.component.getSheets()

    def get_controller(self) -> XController:
        """
        Provides access to the controller which currently controls this model

        Raises:
            MissingInterfaceError: If unable to access controller

        Returns:
            XController | None: Controller for Spreadsheet Document
        """
        model = self.qi(XModel, True)
        return model.getCurrentController()

    def get_active_sheet(self) -> mCalcSheet.CalcSheet:
        """
        Gets the active sheet

        Returns:
            CalcSheet | None: Active Sheet if found; Otherwise, None
        """
        with LoContext(self.lo_inst):
            result = mCalc.Calc.get_active_sheet(self.component)
        return mCalcSheet.CalcSheet(owner=self, sheet=result)

    def get_view(self) -> mCalcSheetView.CalcSheetView:
        """
        Is the main interface of a SpreadsheetView.

        It manages the active sheet within this view.

        The ``com.sun.star.sheet.SpreadsheetView`` service is the spreadsheet's extension
        of the ``com.sun.star.frame.Controller`` service and represents a table editing view
        for a spreadsheet document.

        Returns:
            mCalcSheetView: CalcSheetView
        """
        with LoContext(self.lo_inst):
            view = mCalc.Calc.get_view(self.component)
        return mCalcSheetView.CalcSheetView(owner=self, view=view, lo_inst=self.lo_inst)

    # region set_active_sheet()
    @overload
    def set_active_sheet(self, sheet: XSpreadsheet) -> None:
        """
        Sets the active sheet

        Args:
            sheet (XSpreadsheet): Sheet to set active
        """
        ...

    @overload
    def set_active_sheet(self, sheet: mCalcSheet.CalcSheet) -> None:
        """
        Sets the active sheet

        Args:
            sheet (CalcSheet): Sheet to set active
        """
        ...

    def set_active_sheet(self, sheet: Any) -> None:
        """
        Sets the active sheet

        Args:
            sheet (XSpreadsheet): Sheet to set active

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ACTIVATING` :eventref:`src-docs-sheet-event-activating`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_ACTIVATED` :eventref:`src-docs-sheet-event-activated`

        Note:
            Event arg properties modified on SHEET_ACTIVATING it is reflected in this method.
        """
        with LoContext(self.lo_inst):
            if mInfo.Info.is_instance(sheet, mCalcSheet.CalcSheet):
                mCalc.Calc.set_active_sheet(self.component, sheet.component)
            else:
                mCalc.Calc.set_active_sheet(self.component, sheet)

    # endregion set_active_sheet()

    def set_view_data(self, view_data: str) -> None:
        """
        Restores the view status using the data gotten from a previous call to
        ``get_view_data()``

        Args:
            view_data (str): Data to restore.
        """
        with LoContext(self.lo_inst):
            mCalc.Calc.set_view_data(self.component, view_data)

    def get_view_data(self) -> str:
        """
        Gets a set of data that can be used to restore the current view status at
        later time by using :py:meth:`~Calc.set_view_data`

        Returns:
            str: View Data
        """
        with LoContext(self.lo_inst):
            result = mCalc.Calc.get_view_data(self.component)
        return result

    def get_view_panes(self) -> List[XViewPane] | None:
        """
        represents a pane in a view of a spreadsheet document.

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document

        Raises:
            MissingInterfaceError: if unable access the view pane container

        Returns:
            List[XViewPane] | None: List of XViewPane on success; Otherwise, None

        Note:
            The com.sun.star.sheet.XViewPane interface's getFirstVisibleColumn(), getFirstVisibleRow(),
            setFirstVisibleColumn() and setFirstVisibleRow() methods query and set the start of
            the exposed area. The getVisibleRange() method returns a com.sun.star.table.
            CellRangeAddress struct describing which cells are shown in the pane.
            Columns or rows that are only partly visible at the right or lower edge of the view
            are not included.

        See Also:
            :ref:`ch23_view_states_top_pane`
        """
        with LoContext(self.lo_inst):
            result = mCalc.Calc.get_view_panes(doc=self.component)
        return result

    def get_view_states(self) -> List[mViewState.ViewState] | None:
        """
        Extract the view states for all the sheets from the view data.
        The states are returned as an array of ViewState objects.

        The view data string has the format
        ``100/60/0;0;tw:879;0/4998/0/1/0/218/2/0/0/4988/4998``

        The view state info starts after the third ``;``, the fourth entry.
        The view state for each sheet is separated by ``;``

        Based on a post by user Hanya to:
        `openoffice forum <https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=29195&p=133202&hilit=getViewData#p133202>`_

        Returns:
            None:

        Hint:
            - ``ViewState`` can be imported from ``ooodev.utils.view_state``

        See Also:
            :ref:`ch23_view_states_top_pane`
        """
        with LoContext(self.lo_inst):
            result = mCalc.Calc.get_view_states(doc=self.component)
        return result

    def set_view_states(self, states: Sequence[mViewState.ViewState]) -> None:
        """
        Updates the sheet state part of the view data, which starts as the fourth entry in the view data string.

        Args:
            states (Sequence[ViewState]): Sequence of ViewState objects.

        Returns:
            None:

        Hint:
            - ``ViewState`` can be imported from ``ooodev.utils.view_state``

        See Also:
            :ref:`ch23_view_states_top_pane`.
        """
        with LoContext(self.lo_inst):
            mCalc.Calc.set_view_states(doc=self.component, states=states)

    def split_window(self, cell_name: str) -> None:
        """
        Splits window

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            cell_name (str): Cell to preform split on. e.g. 'C4'

        Returns:
            None:

        See Also:
            :ref:`ch23_splitting_panes`
        """
        with LoContext(self.lo_inst):
            # split window dispatches GoToCell and SplitWindow
            # use context manager to observe global events while dispatch is taking place.
            # this will allow CalcDoc to observe the event.
            # eg: doc.subscribe_event("lo_dispatching", on_dispatching)
            with observe_events(self.event_observer):
                mCalc.Calc.split_window(doc=self.component, cell_name=cell_name)

    @overload
    def set_visible(self) -> None:
        """
        Set window visible.

        Returns:
            None:
        """
        ...

    @overload
    def set_visible(self, visible: bool) -> None:
        """
        Set window visibility.

        Args:
            visible (bool, optional): If ``True`` window is set visible; Otherwise, window is set invisible. Default ``True``

        Returns:
            None:
        """
        ...

    def set_visible(self, visible: bool = True) -> None:
        """
        Set window visibility.

        Args:
            visible (bool, optional): If ``True`` window is set visible; Otherwise, window is set invisible. Default ``True``

        Returns:
            None:
        """
        mGUI.GUI.set_visible(doc=self.component, visible=visible)

    def unfreeze(self) -> None:
        """
        UN-Freezes spreadsheet columns and/or rows

        Returns:
            None:

        See Also:

            - :ref:`ch23_freezing_rows`
            - :py:meth:`~.Calc.freeze`
            - :py:meth:`~.Calc.freeze_rows`
            - :py:meth:`~.Calc.freeze_cols`
        """
        with LoContext(self.lo_inst):
            mCalc.Calc.unfreeze(doc=self.component)

    def zoom_value(self, value: int = 100) -> None:
        """
        Sets the zoom level of the Spreadsheet Document

        Args:
            value (int, optional): Value to set zoom. e.g. 160 set zoom to 160%. Default is ``100``.
        """
        with LoContext(self.lo_inst):
            mCalc.Calc.zoom_value(doc=self.component, value=value)

    def zoom(self, type: ZoomKind = ZoomKind.ZOOM_100_PERCENT) -> None:
        """
        Zooms spreadsheet document to a specific view.

        Args:
            type (ZoomKind, optional): Type of Zoom to set. Default is ``ZoomKind.ZOOM_100_PERCENT``.

        Hint:
            - ``ZoomKind`` can be imported from ``ooodev.utils.kind.zoom_kind``
        """
        with LoContext(self.lo_inst):
            mCalc.Calc.zoom(doc=self.component, type=type)

    # region create_doc()

    @classmethod
    def _on_io_creating_doc(cls, event_args: CancelEventArgs) -> None:
        """
        Event called before document is Created.

        Args:
            event_args (CancelEventArgs): Event data.

        Raises:
            CancelEventError: If event is canceled.
        """
        _Events().trigger(CalcNamedEvent.DOC_CREATING, event_args)

    @classmethod
    def _on_io_created_doc(cls, event_args: EventArgs) -> None:
        """
        Event called after document is Created.

        Args:
            event_args (EventArgs): Event data.
        """
        _Events().trigger(CalcNamedEvent.DOC_CREATED, event_args)

    # endregion create_doc()

    # region open_doc()

    @classmethod
    def _on_io_opening_doc(cls, event_args: CancelEventArgs) -> None:
        """
        Event called before document is Opened.

        Args:
            event_args (CancelEventArgs): Event data.

        Raises:
            CancelEventError: If event is canceled.
        """
        _Events().trigger(CalcNamedEvent.DOC_OPENING, event_args)

    @classmethod
    def _on_io_opened_doc(cls, event_args: EventArgs) -> None:
        """
        Event called after document is Opened.

        Args:
            event_args (EventArgs): Event data.
        """
        _Events().trigger(CalcNamedEvent.DOC_OPENED, event_args)

    # endregion open_doc()

    # region from_current_doc()
    @classmethod
    def _on_from_current_doc_loading(cls, event_args: CancelEventArgs) -> None:
        """
        Event called while from_current_doc loading.

        Args:
            event_args (EventArgs): Event data.

        Returns:
            None:

        Note:
            event_args.event_data is a dictionary and contains the document in a key named 'doc'.
        """
        event_args.event_data["doc_type"] = cls.DOC_TYPE

    @classmethod
    def _on_from_current_doc_loaded(cls, event_args: EventArgs) -> None:
        """
        Event called after from_current_doc is called.

        Args:
            event_args (EventArgs): Event data.

        Returns:
            None:

        Note:
            event_args.event_data is a dictionary and contains the document in a key named 'doc'.
        """
        doc = cast(CalcDoc, event_args.event_data["doc"])
        if doc.DOC_TYPE != cls.DOC_TYPE:
            raise mEx.NotSupportedDocumentError(f"Document '{type(doc).__name__}' is not a Calc document.")

    # endregion from_current_doc()
    # region Properties

    @property
    def sheets(self) -> mCalcSheets.CalcSheets:
        """
        Sheets of Calc Document

        Returns:
            CalcSheets: Calc Sheets

        .. code-block:: python

            # example of setting the value of cell A2 to TEST
            doc.sheets[0]["A2"].set_val("TEST")

        .. versionadded:: 0.17.11
        """
        if self._sheets is None:
            self._sheets = mCalcSheets.CalcSheets(owner=self, sheets=self.component.getSheets(), lo_inst=self.lo_inst)
        return self._sheets

    @property
    def draw_pages(self) -> SpreadsheetDrawPages[CalcDoc]:
        """
        Gets draw pages.

        Returns:
            SpreadsheetDrawPages: Draw Pages
        """
        if self._draw_pages is None:
            supp = self.qi(XDrawPagesSupplier, True)
            draw_pages = supp.getDrawPages()
            self._draw_pages = SpreadsheetDrawPages(owner=self, slides=draw_pages, lo_inst=self.lo_inst)
        return self._draw_pages  # type: ignore

    @property
    def current_controller(self) -> mCalcSheetView.CalcSheetView:
        """
        Gets the current controller.

        Returns:
            CalcSheetView: Current Controller
        """
        if self._current_controller is None:
            self._current_controller = mCalcSheetView.CalcSheetView(owner=self, view=self.component.getCurrentController(), lo_inst=self.lo_inst)  # type: ignore
        return self._current_controller

    # endregion Properties
