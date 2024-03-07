from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sheet import XSpreadsheetView

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

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

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

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
