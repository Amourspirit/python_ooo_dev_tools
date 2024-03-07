from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.chart import XChartData

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.chart import XChartDataChangeEventListener
    from ooodev.utils.type_var import UnoInterface


class ChartDataPartial:
    """
    Partial class for XChartData.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XChartData, interface: UnoInterface | None = XChartData) -> None:
        """
        Constructor

        Args:
            component (XChartData): UNO Component that implements ``com.sun.star.chart.XChartData`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XChartData``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XChartData
    def add_chart_data_change_event_listener(self, listener: XChartDataChangeEventListener) -> None:
        """
        allows a component supporting the XChartDataChangeEventListener interface to register as listener.

        The component will be notified with a ChartDataChangeEvent every time the chart's data changes.
        """
        self.__component.addChartDataChangeEventListener(listener)

    def get_not_a_number(self) -> float:
        """
        In IEEE arithmetic format it is one of the NaN values, so there are no conflicts with existing numeric values.
        """
        return self.__component.getNotANumber()

    def is_not_a_number(self, number: float) -> bool:
        """
        checks whether the value given is equal to the indicator value for a missing value.

        In IEEE arithmetic format it is one of the NaN values, so there are no conflicts with existing numeric values.

        Always use this method to check, if a value is not a number. If you compare the value returned by XChartData.getNotANumber() to another double value using the = operator, you may not get the desired result!
        """
        return self.__component.isNotANumber(number)

    def remove_chart_data_change_event_listener(self, listener: XChartDataChangeEventListener) -> None:
        """
        Removes a previously registered listener.
        """
        self.__component.removeChartDataChangeEventListener(listener)

    # endregion XChartData
