from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno
import contextlib

from com.sun.star.chart2 import XChartTypeContainer

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartType
    from ooodev.utils.type_var import UnoInterface


# This class may be implement by class that dont support the interface.
# For this reason all methods are wrapped with contextlib.suppress(AttributeError)


class ChartTypeContainerPartial:
    """
    Partial class for XChartTypeContainer.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XChartTypeContainer, interface: UnoInterface | None = XChartTypeContainer) -> None:
        """
        Constructor

        Args:
            component (XChartTypeContainer): UNO Component that implements ``com.sun.star.chart2.XChartTypeContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XChartTypeContainer``.
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

    # region XChartTypeContainer
    def add_chart_type(self, chart_types: XChartType) -> None:
        """
        Add a chart type to the chart type container

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        with contextlib.suppress(AttributeError):
            self.__component.addChartType(chart_types)

    def get_chart_types(self) -> Tuple[XChartType, ...]:
        """
        Gets all chart types
        """
        with contextlib.suppress(AttributeError):
            return self.__component.getChartTypes()
        return ()

    def remove_chart_type(self, chart_types: XChartType) -> None:
        """
        Removes one data series from the chart type container.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        with contextlib.suppress(AttributeError):
            self.__component.removeChartType(chart_types)

    def set_chart_types(self, *chart_types: XChartType) -> None:
        """
        Set all chart types

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        with contextlib.suppress(AttributeError):
            self.__component.setChartTypes(chart_types)

    # endregion XChartType
