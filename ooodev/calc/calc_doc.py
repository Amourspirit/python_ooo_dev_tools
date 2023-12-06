from __future__ import annotations
from typing import Any, List, Tuple, overload, Sequence, TYPE_CHECKING
import uno

from com.sun.star.frame import XModel
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.sheet import XSpreadsheets


if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.frame import XController
    from com.sun.star.sheet import XViewPane
    from com.sun.star.style import XStyle
    from com.sun.star.table import CellAddress
    from com.sun.star.table import XCellRange
    from ooo.dyn.sheet.general_function import GeneralFunction
    from ooo.dyn.table.cell_range_address import CellRangeAddress

    from ooodev.utils.kind.zoom_kind import ZoomKind
else:
    CellRangeAddress = object

from ooodev.adapter.sheet.spreadsheet_document_comp import SpreadsheetDocumentComp
from ooodev.events.args.calc.sheet_args import SheetArgs
from ooodev.events.args.calc.sheet_cancel_args import SheetCancelArgs
from ooodev.events.calc_named_event import CalcNamedEvent
from ooodev.events.event_singleton import _Events
from ooodev.office import calc as mCalc
from ooodev.utils import gui as mGUI
from ooodev.utils import lo as mLo
from ooodev.utils import lo as mLo
from ooodev.utils import view_state as mViewState
from ooodev.utils.data_type import range_obj as mRngObj
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.type_var import PathOrStr
from . import calc_sheet as mCalcSheet
from . import calc_sheet_view as mCalcSheetView


class CalcDoc(SpreadsheetDocumentComp, QiPartial, PropPartial):
    def __init__(self, doc: XSpreadsheetDocument) -> None:
        SpreadsheetDocumentComp.__init__(self, doc)  # type: ignore
        QiPartial.__init__(self, component=doc, lo_inst=mLo.Lo.current_lo)
        PropPartial.__init__(self, component=doc, lo_inst=mLo.Lo.current_lo)

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
        return mCalc.Calc.create_cell_style(doc=self.component, style_name=style_name)

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
        return mCalc.Calc.call_fun(func_name, *args)

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
        """
        return mCalc.Calc.compute_function(fn=fn, cell_range=cell_range)

    def get_function_names(self) -> List[str] | None:
        """
        Get a list of all function names.

        Returns:
            List[str] | None: List of function names if found; Otherwise, ``None``.
        """
        return mCalc.Calc.get_function_names()

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
        return mCalc.Calc.find_function(*args, **kwargs)

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
            deliver_ownership (bool): If ``True`` delegates the ownership of this closing object to
                anyone which throw the CloseVetoException. Default is ``False``.

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
        mLo.Lo.close_doc(doc=self.component, deliver_ownership=deliver_ownership)

    # endregion close_doc()

    def save_doc(self, fnm: PathOrStr) -> bool:
        """
        Saves text document

        Args:
            fnm (PathOrStr): Path to save as

        Raises:
            MissingInterfaceError: If doc does not implement XComponent interface

        Returns:
            bool: True if doc is saved; Otherwise, False

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.DOC_SAVING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.DOC_SAVED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing ``text_doc`` and ``fnm``.

        Attention:
            :py:meth:`Lo.save_doc <.utils.lo.Lo.save_doc>` method is called along with any of its events.
        """
        return mCalc.Calc.save_doc(doc=self.component, fnm=fnm)

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
        return mCalc.Calc.get_selected_addr(doc=self.component)

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
        return mCalc.Calc.get_selected_cell_addr(self.component)

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

        Raises:
            Exception: If spreadsheet is not found
            CancelEventError: If SHEET_GETTING event is canceled

        Returns:
            CalcSheet: Spreadsheet at index.
        """
        ...

    @overload
    def get_sheet(self, idx: int) -> mCalcSheet.CalcSheet:
        """
        Gets a sheet of spreadsheet document

        Args:
            idx (int, optional): Zero based index of spreadsheet. Defaults to ``0``

        Raises:
            Exception: If spreadsheet is not found
            CancelEventError: If SHEET_GETTING event is canceled

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

        Raises:
            Exception: If spreadsheet is not found
            CancelEventError: If SHEET_GETTING event is canceled

        Returns:
            CalcSheet: Spreadsheet at index.
        """
        ...

    def get_sheet(self, *args, **kwargs) -> mCalcSheet.CalcSheet:
        """
        Gets a sheet of spreadsheet document

        Args:
            idx (int, optional): Zero based index of spreadsheet. Defaults to ``0``
            sheet_name (str, optional): Name of spreadsheet

        Raises:
            Exception: If spreadsheet is not found
            CancelEventError: If SHEET_GETTING event is canceled

        Returns:
            CalcSheet: Spreadsheet at index.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_GETTING` :eventref:`src-docs-sheet-event-getting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_GET` :eventref:`src-docs-sheet-event-get`

        Note:
            For Event args, if ``index`` is available then ``name`` is ``None`` and if ``sheet_name`` is available then ``index`` is ``None``.
        """
        sheet = mCalc.Calc.get_sheet(self.component, *args, **kwargs)
        return mCalcSheet.CalcSheet(owner=self, sheet=sheet)

    # endregion get_sheet()

    def insert_sheet(self, name: str, idx: int) -> mCalcSheet.CalcSheet:
        """
        Inserts a spreadsheet into document.

        Args:
            name (str): Name of sheet to insert
            idx (int): zero-based index position of the sheet to insert

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
        # sourcery skip: raise-specific-error
        sheet = mCalc.Calc.insert_sheet(doc=self.component, name=name, idx=idx)
        return mCalcSheet.CalcSheet(owner=self, sheet=sheet)

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
            return mCalc.Calc.remove_sheet(doc=self.component, idx=kwargs["idx"])
        if kwargs:
            raise TypeError("remove_sheet() got an unexpected keyword argument")
        if count != 1:
            raise TypeError("remove_sheet() got an invalid number of arguments")
        arg = args[0]
        if isinstance(arg, int):
            return mCalc.Calc.remove_sheet(doc=self.component, idx=arg)
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
        cargs = SheetCancelArgs(self.move_sheet.__qualname__)
        cargs.doc = self.component
        cargs.name = name
        cargs.index = idx
        _Events().trigger(CalcNamedEvent.SHEET_MOVING, cargs)
        if cargs.cancel:
            return False
        name = cargs.name
        idx = cargs.index
        sheets = cargs.doc.getSheets()
        num_sheets = len(sheets.getElementNames())
        result = False
        if idx < 0 or idx >= num_sheets:
            mLo.Lo.print(f"Index {idx} is out of range.")
        else:
            sheets.moveByName(name, idx)
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
        model = mLo.Lo.qi(XModel, self.component, True)
        return model.getCurrentController()

    def get_active_sheet(self) -> mCalcSheet.CalcSheet:
        """
        Gets the active sheet

        Returns:
            CalcSheet | None: Active Sheet if found; Otherwise, None
        """
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
        view = mCalc.Calc.get_view(self.component)
        return mCalcSheetView.CalcSheetView(self, view)

    def set_active_sheet(self, sheet: XSpreadsheet) -> None:
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
        mCalc.Calc.set_active_sheet(self.component, sheet)

    def set_view_data(self, view_data: str) -> None:
        """
        Restores the view status using the data gotten from a previous call to
        ``get_view_data()``

        Args:
            view_data (str): Data to restore.
        """
        mCalc.Calc.set_view_data(self.component, view_data)

    def get_view_data(self) -> str:
        """
        Gets a set of data that can be used to restore the current view status at
        later time by using :py:meth:`~Calc.set_view_data`

        Returns:
            str: View Data
        """
        return mCalc.Calc.get_view_data(self.component)

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
        return mCalc.Calc.get_view_panes(doc=self.component)

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

        See Also:
            :ref:`ch23_view_states_top_pane`
        """
        return mCalc.Calc.get_view_states(doc=self.component)

    def set_view_states(self, states: Sequence[mViewState.ViewState]) -> None:
        """
        Updates the sheet state part of the view data, which starts as the fourth entry in the view data string.

        Args:
            states (Sequence[ViewState]): Sequence of ViewState objects.

        Returns:
            None:

        See Also:
            :ref:`ch23_view_states_top_pane`
        """
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
        mCalc.Calc.unfreeze(doc=self.component)

    def zoom_value(self, value: int) -> None:
        """
        Sets the zoom level of the Spreadsheet Document

        Args:
            value (int): Value to set zoom. e.g. 160 set zoom to 160%
        """
        mCalc.Calc.zoom_value(doc=self.component, value=value)

    def zoom(self, type: ZoomKind) -> None:
        """
        Zooms spreadsheet document to a specific view.

        Args:
            doc (XSpreadsheetDocument): Spreadsheet Document
            type (GUI.ZoomEnum): Type of Zoom to set.
        """
        mCalc.Calc.zoom(doc=self.component, type=type)
