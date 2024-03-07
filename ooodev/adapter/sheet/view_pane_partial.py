from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sheet import XViewPane

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.table import CellRangeAddress  # Struct


class ViewPanePartial:
    """
    Partial Class for XViewPane.

    .. versionadded:: 0.20.0
    """

    def __init__(self, component: XViewPane, interface: UnoInterface | None = XViewPane) -> None:
        """
        Constructor

        Args:
            component (XViewPane): UNO Component that implements ``com.sun.star.sheet.XViewPane``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XViewPane``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XViewPane
    def get_first_visible_column(self) -> int:
        """
        Gets the first column that is visible in the pane.
        """
        return self.__component.getFirstVisibleColumn()

    def get_first_visible_row(self) -> int:
        """
        Gets the first row that is visible in the pane.
        """
        return self.__component.getFirstVisibleRow()

    def get_visible_range(self) -> CellRangeAddress:
        """
        Gets the address of the cell range that consists of the cells which are visible in the pane.
        """
        return self.__component.getVisibleRange()

    def set_first_visible_column(self, column: int) -> None:
        """
        Sets the first column that is visible in the pane.
        """
        self.__component.setFirstVisibleColumn(column)

    def set_first_visible_row(self, row: int) -> None:
        """
        Sets the first row that is visible in the pane.
        """
        self.__component.setFirstVisibleRow(row)

    # endregion XViewPane
