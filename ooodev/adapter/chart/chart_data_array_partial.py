from __future__ import annotations
from typing import TYPE_CHECKING, Tuple

import uno
from com.sun.star.chart import XChartDataArray

from ooodev.adapter.chart.chart_data_partial import ChartDataPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ChartDataArrayPartial(ChartDataPartial):
    """
    Partial class for XChartDataArray.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XChartDataArray, interface: UnoInterface | None = XChartDataArray) -> None:
        """
        Constructor

        Args:
            component (XChartDataArray ): UNO Component that implements ``com.sun.star.chart.XChartDataArray`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XChartDataArray``.
        """

        ChartDataPartial.__init__(self, component=component, interface=interface)
        self.__component = component

    # region XChartDataArray
    def get_column_descriptions(self) -> Tuple[str, ...]:
        """
        Gets the description texts for all columns.
        """
        return self.__component.getColumnDescriptions()

    def get_data(self) -> Tuple[Tuple[float, ...], ...]:
        """
        Gets the numerical data as a nested sequence of values.
        """
        return self.__component.getData()

    def get_row_descriptions(self) -> Tuple[str, ...]:
        """
        retrieves the description texts for all rows.
        """
        return self.__component.getRowDescriptions()

    def set_column_descriptions(self, *descriptions: str) -> None:
        """
        Sets the description texts for all columns.
        """
        self.__component.setColumnDescriptions(descriptions)

    def set_data(self, data: Tuple[Tuple[float, ...], ...]) -> None:
        """
        Sets the chart data as an array of numbers.
        """
        self.__component.setData(data)

    def set_row_descriptions(self, *descriptions: str) -> None:
        """
        sets the description texts for all rows.
        """
        self.__component.setRowDescriptions(descriptions)

    # endregion XChartDataArray
