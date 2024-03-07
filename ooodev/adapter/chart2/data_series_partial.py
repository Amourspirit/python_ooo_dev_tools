from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.chart2 import XDataSeries

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import XPropertySet
    from ooodev.utils.type_var import UnoInterface


class DataSeriesPartial:
    """
    Partial class for XDataSeries.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDataSeries, interface: UnoInterface | None = XDataSeries) -> None:
        """
        Constructor

        Args:
            component (XDataSeries): UNO Component that implements ``com.sun.star.chart2.XDataSeries`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDataSeries``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDataSeries
    def get_data_point_by_index(self, idx: int) -> XPropertySet:
        """

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        return self.__component.getDataPointByIndex(idx)

    def reset_all_data_points(self) -> None:
        """
        all data point formatting are cleared
        """
        self.__component.resetAllDataPoints()

    def reset_data_point(self, idx: int) -> None:
        """
        the formatting of the specified data point is cleared
        """
        self.__component.resetDataPoint(idx)

    # endregion XDataSeries
