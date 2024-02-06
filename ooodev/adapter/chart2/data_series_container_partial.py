from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.chart2 import XDataSeriesContainer

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.chart2 import XDataSeries


class DataSeriesContainerPartial:
    """
    Partial class for XDataSeriesContainer.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDataSeriesContainer, interface: UnoInterface | None = XDataSeriesContainer) -> None:
        """
        Constructor

        Args:
            component (XDataSeriesContainer): UNO Component that implements ``com.sun.star.chart2.XDataSeriesContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDataSeriesContainer``.
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

    # region XDataSeriesContainer
    def add_data_series(self, data_series: XDataSeries) -> None:
        """
        Add a data series to the data series container

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.addDataSeries(data_series)

    def get_data_series(self) -> Tuple[XDataSeries, ...]:
        """
        retrieve all data series
        """
        ...

    def remove_data_series(self, data_series: XDataSeries) -> None:
        """
        removes one data series from the data series container.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        self.__component.removeDataSeries(data_series)

    def set_data_series(self, data_series: Tuple[XDataSeries, ...]) -> None:
        """
        Set all data series.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.setDataSeries(data_series)

    # endregion XDataSeriesContainer
