from __future__ import annotations
from typing import TYPE_CHECKING

from com.sun.star.sheet import XSpreadsheets
from ooodev.adapter.container.name_container_partial import NameContainerPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class SpreadsheetsPartial(NameContainerPartial):
    """
    Partial Class for XSpreadsheets.
    """

    def __init__(self, component: XSpreadsheets, interface: UnoInterface | None = XSpreadsheets) -> None:
        """
        Constructor

        Args:
            component (XSpreadsheets): UNO Component that implements ``com.sun.star.sheet.XSpreadsheets``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSpreadsheets``.
        """
        NameContainerPartial.__init__(self, component, interface)
        self.__component = component

    # region XSpreadsheets
    def copy_by_name(self, name: str, copy: str, idx: int) -> None:
        """
        Copies the sheet with the specified name.

        Args:
            name (str): The name of the sheet to be copied.
            copy (str): The name of the copy of the spreadsheet.
            idx (int, optional): The index of the copy in the collection.

        """
        self.__component.copyByName(name, copy, idx)

    def insert_new_by_name(self, name: str, idx: int) -> None:
        """
        Inserts a new sheet with the specified name.

        Args:
            name (str): The name of the sheet to be inserted.
            idx (int, optional): The index of the new sheet.
        """
        self.__component.insertNewByName(name, idx)

    def move_by_name(self, name: str, idx: int) -> None:
        """
        Moves the sheet with the specified name.

        Args:
            name (str): The name of the sheet to be moved.
            idx (int): The new index of the sheet.
        """
        self.__component.moveByName(name, idx)

    # endregion XSpreadsheets
