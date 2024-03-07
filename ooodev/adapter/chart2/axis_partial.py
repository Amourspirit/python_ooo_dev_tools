from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.chart2 import XAxis

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import XPropertySet
    from com.sun.star.chart2 import ScaleData  # Struct
    from ooodev.utils.type_var import UnoInterface


class AxisPartial:
    """
    Partial class for XAxis.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XAxis, interface: UnoInterface | None = XAxis) -> None:
        """
        Constructor

        Args:
            component (XAxis): UNO Component that implements ``com.sun.star.chart2.XAxis`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XAxis``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XAxis

    def get_grid_properties(self) -> XPropertySet:
        """
        Gets the returned property set must support the service GridProperties
        """
        return self.__component.getGridProperties()

    def get_scale_data(self) -> ScaleData:
        """
        Gets the scale data.
        """
        return self.__component.getScaleData()

    def get_sub_grid_properties(self) -> Tuple[XPropertySet, ...]:
        """
        the returned property sets must support the service GridProperties

        If you do not want to render certain a sub-grid, in the corresponding XPropertySet the property GridProperties.Show must be FALSE.
        """
        return self.__component.getSubGridProperties()

    def get_sub_tick_properties(self) -> Tuple[XPropertySet, ...]:
        """
        the returned property sets must support the service TickProperties
        """
        return self.__component.getSubTickProperties()

    def set_scale_data(self, scale: ScaleData) -> None:
        """
        Sets the scale data.
        """
        return self.__component.setScaleData(scale)

    # endregion XAxis
