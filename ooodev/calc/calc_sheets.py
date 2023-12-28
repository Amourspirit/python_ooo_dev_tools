from __future__ import annotations
from typing import overload, TYPE_CHECKING, Tuple

import uno
from . import calc_sheet as mCalcSheet
from ooodev.office import calc as mCalc
from ooodev.adapter.sheet.spreadsheets_comp import SpreadsheetsComp
from ooodev.utils import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.adapter.sheet.cell_range_access_partial import CellRangeAccessPartial

if TYPE_CHECKING:
    from .calc_doc import CalcDoc
    from com.sun.star.sheet import XSpreadsheet
    from com.sun.star.sheet import XSpreadsheets


class CalcSheets(SpreadsheetsComp, CellRangeAccessPartial, QiPartial):
    """
    Class for managing Calc Sheets.

    This class is Enumerable and returns ``CalcSheet`` instance on iteration.

    .. code-block:: python

        for sheet in doc.sheets:
            sheet["A1"].set_val("test")
            assert sheet["A1"].get_val() == "test"

    This class also as index access and returns ``CalcSheet`` instance.

    .. code-block:: python

        sheet = doc.sheets["Sheet1"]
        # or set the value of cell A2 to TEST
        doc.sheets[0]["A2"].set_val("TEST")

        # get the last sheet of the document
        last_sheet = doc.sheets[-1]

        # get the second last sheet of the document
        second_last_sheet = doc.sheets[-2]

        # get the number of sheets
        num_sheets = len(doc.sheets)

    .. versionchanged:: 0.17.13
        - Added negative index access.
        - Added ``__len__`` method.

    .. versionadded:: 0.17.11
    """

    def __init__(self, owner: CalcDoc, sheets: XSpreadsheets) -> None:
        """
        Constructor

        Args:
            owner (CalcDoc): Owner Document
            sheet (XSpreadsheet): Sheet instance.
        """
        self.__owner = owner
        SpreadsheetsComp.__init__(self, sheets)  # type: ignore
        CellRangeAccessPartial.__init__(self, component=sheets, interface=None)  # type: ignore
        QiPartial.__init__(self, component=sheets, lo_inst=mLo.Lo.current_lo)

    def __next__(self) -> mCalcSheet.CalcSheet:
        return mCalcSheet.CalcSheet(owner=self.__owner, sheet=super().__next__())

    def __getitem__(self, index: str | int) -> mCalcSheet.CalcSheet:
        if isinstance(index, int):
            if index < 0:
                index = len(self) + index
                if index < 0:
                    raise IndexError("list index out of range")
        return self.get_sheet(index)

    def __len__(self) -> int:
        return self.component.getCount()

    def get_active_sheet(self) -> mCalcSheet.CalcSheet:
        """
        Gets the active sheet

        Returns:
            CalcSheet | None: Active Sheet if found; Otherwise, None
        """
        result = mCalc.Calc.get_active_sheet(self.calc_doc.component)
        return mCalcSheet.CalcSheet(owner=self.calc_doc, sheet=result)

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
            idx (int, optional): Zero based index of spreadsheet. Defaults to ``0``

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
            idx (int, optional): Zero based index of spreadsheet. Defaults to ``0``
            sheet_name (str, optional): Name of spreadsheet

        Raises:
            Exception: If spreadsheet is not found
            CancelEventError: If ``SHEET_GETTING`` event is canceled

        Returns:
            CalcSheet: Spreadsheet at index.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_GETTING` :eventref:`src-docs-sheet-event-getting`
                - :py:attr:`~.events.calc_named_event.CalcNamedEvent.SHEET_GET` :eventref:`src-docs-sheet-event-get`

        Note:
            For Event args, if ``index`` is available then ``name`` is ``None`` and if ``sheet_name`` is available then ``index`` is ``None``.
        """
        sheet = mCalc.Calc.get_sheet(self.calc_doc.component, *args, **kwargs)
        return mCalcSheet.CalcSheet(owner=self.calc_doc, sheet=sheet)

    # endregion get_sheet()

    def get_sheet_names(self) -> Tuple[str, ...]:
        """
        Gets names of all existing spreadsheets in the spreadsheet document.

        Returns:
            Tuple[str, ...]: Tuple of sheet names.
        """
        return self.calc_doc.get_sheet_names()

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

    def set_active_sheet(self, sheet: XSpreadsheet | mCalcSheet.CalcSheet) -> None:
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
        self.calc_doc.set_active_sheet(sheet)

    # endregion set_active_sheet()

    def insert_sheet(self, name: str, idx: int) -> mCalcSheet.CalcSheet:
        """
        Inserts a spreadsheet into document.

        Args:
            name (str): Name of sheet to insert
            idx (int): zero-based index position of the sheet to insert.
                Can be a negative value to insert from the end of the list.

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
        return self.calc_doc.insert_sheet(name, idx)

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
                Can be a negative value to insert from the end of the list.

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
        return self.calc_doc.remove_sheet(*args, **kwargs)

    # endregion remove_sheet()

    # region Properties
    @property
    def calc_doc(self) -> CalcDoc:
        """
        Returns:
            CalcDoc: Calc doc
        """
        return self.__owner

    # endregion Properties
