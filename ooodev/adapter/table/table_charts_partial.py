from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
import uno

from com.sun.star.table import XTableCharts
from ooodev.adapter.container.name_access_partial import NameAccessPartial

if TYPE_CHECKING:
    from com.sun.star.awt import Rectangle  # Struct
    from com.sun.star.table import CellRangeAddress  # Struct
    from ooodev.utils.type_var import UnoInterface


class TableChartsPartial(NameAccessPartial):
    """
    Partial Class for XTableCharts.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTableCharts, interface: UnoInterface | None = XTableCharts) -> None:
        """
        Constructor

        Args:
            component (XTableCharts): UNO Component that implements ``com.sun.star.container.XTableCharts`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTableCharts``.
        """
        NameAccessPartial.__init__(self, component, interface)
        self.__component = component

    # region XTableCharts
    def add_new_by_name(
        self, name: str, rect: Rectangle, ranges: Tuple[CellRangeAddress, ...], column_headers: bool, row_headers: bool
    ) -> None:
        """
        Creates a chart and adds it to the collection.
        """

        self.__component.addNewByName(name, rect, ranges, column_headers, row_headers)

    def remove_by_name(self, name: str) -> None:
        """
        Removes a chart from the collection.
        """
        self.__component.removeByName(name)

    # endregion XTableCharts
