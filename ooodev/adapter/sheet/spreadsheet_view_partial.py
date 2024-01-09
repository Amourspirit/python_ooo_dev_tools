from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sheet import XSpreadsheetView

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.sheet import XSpreadsheet
    from ooodev.utils.type_var import UnoInterface


class SpreadsheetViewPartial:
    """
    Partial Class for XSpreadsheetView.

    .. versionadded:: 0.20.0
    """

    def __init__(self, component: XSpreadsheetView, interface: UnoInterface | None = XSpreadsheetView) -> None:
        """
        Constructor

        Args:
            component (XSpreadsheetView): UNO Component that implements ``com.sun.star.sheet.XSpreadsheetView``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSpreadsheetView``.
        """
        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XSpreadsheetView
    def get_active_sheet(self) -> XSpreadsheet:
        """
        Gets the active sheet of the spreadsheet document.
        """
        return self.__component.getActiveSheet()

    def set_active_sheet(self, sheet: XSpreadsheet) -> None:
        """
        Sets the active sheet of the spreadsheet document.
        """
        self.__component.setActiveSheet(sheet)

    # endregion XSpreadsheetView
