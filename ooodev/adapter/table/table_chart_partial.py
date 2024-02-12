from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.table import XTableChart

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.table import CellRangeAddress  # Struct
    from ooodev.utils.type_var import UnoInterface


class TableChartPartial:
    """
    Partial Class for XTableChart.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTableChart, interface: UnoInterface | None = XTableChart) -> None:
        """
        Constructor

        Args:
            component (XTableChart): UNO Component that implements ``com.sun.star.container.XTableChart`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTableChart``.
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

    # region XTableChart
    def get_has_column_headers(self) -> bool:
        """
        Returns, whether the cells of the topmost row of the source data are interpreted as column headers.
        """
        return self.__component.getHasColumnHeaders()

    def get_has_row_headers(self) -> bool:
        """
        Returns, whether the cells of the leftmost column of the source data are interpreted as row headers.
        """
        return self.__component.getHasRowHeaders()

    def get_ranges(self) -> Tuple[CellRangeAddress, ...]:
        """
        Returns the cell ranges that contain the data for the chart.
        """
        return self.__component.getRanges()

    def set_has_column_headers(self, has_column_headers: bool) -> None:
        """
        Specifies whether the cells of the topmost row of the source data are interpreted as column headers.
        """
        self.__component.setHasColumnHeaders(has_column_headers)

    def set_has_row_headers(self, has_row_headers: bool) -> None:
        """
        Specifies whether the cells of the leftmost column of the source data are interpreted as row headers.
        """
        self.__component.setHasRowHeaders(has_row_headers)

    def set_ranges(self, ranges: Tuple[CellRangeAddress, ...]) -> None:
        """
        Sets the cell ranges that contain the data for the chart.
        """
        self.__component.setRanges(ranges)

    # endregion XTableChart
