from __future__ import annotations
from typing import cast, overload, TYPE_CHECKING, Tuple
import uno
from com.sun.star.sheet import XSpreadsheet

from ooodev.adapter.container.element_index_partial import ElementIndexPartial
from ooodev.adapter.container.name_replace_partial import NameReplacePartial
from ooodev.adapter.sheet.cell_range_access_partial import CellRangeAccessPartial
from ooodev.adapter.sheet.spreadsheets_comp import SpreadsheetsComp
from ooodev.exceptions import ex as mEx
from ooodev.office import calc as mCalc
from ooodev.utils import gen_util as mGenUtil
from ooodev.utils import info as mInfo
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc import calc_sheet as mCalcSheet

if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheets
    from com.sun.star.sheet import Spreadsheet  # service
    from ooodev.calc.calc_doc import CalcDoc


class CalcSheets(
    LoInstPropsPartial,
    SpreadsheetsComp,
    CellRangeAccessPartial,
    NameReplacePartial["Spreadsheet"],
    QiPartial,
    ServicePartial,
    TheDictionaryPartial,
    ElementIndexPartial,
    CalcDocPropPartial,
):
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

    def __init__(self, owner: CalcDoc, sheets: XSpreadsheets, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (CalcDoc): Owner Document
            sheet (XSpreadsheet): Sheet instance.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        SpreadsheetsComp.__init__(self, sheets)  # type: ignore
        CellRangeAccessPartial.__init__(self, component=sheets, interface=None)  # type: ignore
        NameReplacePartial.__init__(self, component=sheets, interface=None)  # type: ignore
        QiPartial.__init__(self, component=sheets, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=sheets, lo_inst=self.lo_inst)
        TheDictionaryPartial.__init__(self)
        ElementIndexPartial.__init__(self, component=self)  # type: ignore
        CalcDocPropPartial.__init__(self, obj=owner)

    def __next__(self) -> mCalcSheet.CalcSheet:
        """
        Gets the next sheet.

        Returns:
            CalcSheet: The next sheet.
        """
        return mCalcSheet.CalcSheet(owner=self._owner, sheet=super().__next__(), lo_inst=self.lo_inst)

    def __getitem__(self, key: str | int) -> mCalcSheet.CalcSheet:
        """
        Gets the sheet at the specified index or name.

        This is short hand for ``get_by_index()`` or ``get_by_name()``.

        Args:
            key (key, str, int): The index or name of the sheet. When getting by index can be a negative value to get from the end.

        Returns:
            CalcSheet: The sheet with the specified index or name.

        See Also:
            - :py:meth:`~ooodev.calc.CalcSheets.get_by_index`
            - :py:meth:`~ooodev.calc.CalcSheets.get_by_name`
        """
        if isinstance(key, int):
            return self.get_by_index(key)
        return self.get_by_name(key)

    def __len__(self) -> int:
        """
        Gets the number of sheets in the document.

        Returns:
            int: Number of sheet in the document.
        """
        return self.component.getCount()

    def __delitem__(self, _item: int | str | mCalcSheet.CalcSheet) -> None:
        """
        Removes a sheet from the document.

        Args:
            _item (int | str, CalcSheet): Index, name, or sheet to remove.

        Raises:
            TypeError: If the item is not a supported type.
        """
        # using remove_sheet here instead of remove_by_name. This will force Calc module event to be fired.
        if isinstance(_item, (int, str)):
            self.remove_sheet(_item)
        elif isinstance(_item, mCalcSheet.CalcSheet):
            self.remove_sheet(_item.name)
        else:
            raise TypeError(f"Invalid type for __delitem__: {type(_item)}")

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
        count = len(self)
        return mGenUtil.Util.get_index(idx, count, allow_greater)

    # region XSpreadsheets Overrides
    def copy_by_name(self, name: str, copy: str, idx: int) -> None:
        """
        Copies the sheet with the specified name.

        Args:
            name (str): The name of the sheet to be copied.
            copy (str): The name of the copy of the spreadsheet.
            idx (int, optional): The index of the copy in the collection.
                ``idx`` can be a negative value to index from the end of the collection.

        """
        idx = self._get_index(idx)
        super().copy_by_name(name, copy, idx)

    def insert_new_by_name(self, name: str, idx: int) -> None:
        """
        Inserts a new sheet with the specified name.

        Args:
            name (str): The name of the sheet to be inserted.
            idx (int, optional): The index of the new sheet.
                ``idx`` can be a negative value to index from the end of the collection.
        """
        idx = self._get_index(idx=idx, allow_greater=True)
        super().insert_new_by_name(name, idx)

    def move_by_name(self, name: str, idx: int) -> None:
        """
        Moves the sheet with the specified name.

        Args:
            name (str): The name of the sheet to be moved.
            idx (int): The new index of the sheet.
                ``idx`` can be a negative value to index from the end of the collection.
        """
        idx = self._get_index(idx)
        super().move_by_name(name, idx)

    # endregion XSpreadsheets Overrides

    # region XIndexAccess overrides

    def get_by_index(self, idx: int) -> mCalcSheet.CalcSheet:
        """
        Gets the element at the specified index.

        Args:
            idx (int): The Zero-based index of the element. Idx can be a negative value to index from the end of the list.
                For example, -1 will return the last element.

        Returns:
            CalcSheet: The element at the specified index.
        """
        index = self._get_index(idx)
        result = super().get_by_index(index)
        return mCalcSheet.CalcSheet(owner=self.calc_doc, sheet=result, lo_inst=self.lo_inst)

    # endregion XIndexAccess overrides

    # region XNameAccess overrides

    def get_by_name(self, name: str) -> mCalcSheet.CalcSheet:
        """
        Gets the element with the specified name.

        Args:
            name (str): The name of the element.

        Raises:
            MissingNameError: If sheet is not found.

        Returns:
            CalcSheet: The element with the specified name.
        """
        if not self.has_by_name(name):
            raise mEx.MissingNameError(f"Unable to find sheet with name '{name}'")
        result = super().get_by_name(name)
        return mCalcSheet.CalcSheet(owner=self.calc_doc, sheet=result, lo_inst=self.lo_inst)

    # endregion XNameAccess overrides

    def get_active_sheet(self) -> mCalcSheet.CalcSheet:
        """
        Gets the active sheet

        Returns:
            CalcSheet | None: Active Sheet if found; Otherwise, None
        """
        with LoContext(self.lo_inst):
            result = mCalc.Calc.get_active_sheet(self.calc_doc.component)
        return mCalcSheet.CalcSheet(owner=self.calc_doc, sheet=result, lo_inst=self.lo_inst)

    # region    get_sheet()
    @overload
    def get_sheet(self) -> mCalcSheet.CalcSheet:
        """
        Gets the active sheet of spreadsheet document.

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

    @overload
    def get_sheet(self, sheet: XSpreadsheet) -> mCalcSheet.CalcSheet:
        """
        Gets a sheet of spreadsheet document

        Args:
            sheet (XSpreadsheet): Sheet to get as CalcSheet.

        Returns:
            CalcSheet: Spreadsheet from sheet.
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

        .. versionchanged:: 0.20.0
            - Added support for ``XSpreadsheet``.
        """

        args_values = [arg for arg in args]
        args_values.extend(kwargs.values())
        arg_len = len(args_values)
        if arg_len == 0:
            return self.get_active_sheet()
        arg1 = args_values[0]
        if arg_len == 1:
            if mInfo.Info.is_instance(arg1, XSpreadsheet):
                return mCalcSheet.CalcSheet(owner=self.calc_doc, sheet=arg1, lo_inst=self.lo_inst)
            if isinstance(arg1, int):
                return self.get_by_index(arg1)
            return self.get_by_name(cast(str, arg1))
        raise TypeError("get_sheet() got an invalid number of arguments")

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
        return self.calc_doc.insert_sheet(name, idx)

    # endregion insert_sheet

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
