from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.chart2 import XCoordinateSystem

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.chart2 import XAxis
    from ooodev.utils.type_var import UnoInterface


class CoordinateSystemPartial:
    """
    Partial class for XCoordinateSystem.
    """

    def __init__(self, component: XCoordinateSystem, interface: UnoInterface | None = XCoordinateSystem) -> None:
        """
        Constructor

        Args:
            component (XCoordinateSystem): UNO Component that implements ``com.sun.star.chart2.XCoordinateSystem`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCoordinateSystem``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XCoordinateSystem
    def get_axis_by_dimension(self, dimension: int, idx: int) -> XAxis:
        """
        The dimension says whether it is a x, y or z axis.

        The index indicates whether it is a primary or a secondary axis or maybe more in future. Use nIndex == 0 for a primary axis. An empty Reference will be returned if the given nDimension and nIndex are in the valid range but no axis is set for those values. An IndexOutOfBoundsException will be thrown if nDimension is lower than 0 or greater than the value returned by getDimension() and/or if nIndex is lower 0 or greater than the value returned by getMaxAxisIndexByDimension(nDimension).

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        return self.__component.getAxisByDimension(dimension, idx)

    def get_coordinate_system_type(self) -> str:
        """
        Gets the type of coordinate system (e.g.Cartesian, polar ...)
        """
        return self.__component.getCoordinateSystemType()

    def get_dimension(self) -> int:
        """
        the dimension of the coordinate-system.
        """
        return self.__component.getDimension()

    def get_maximum_axis_index_by_dimension(self, dimension: int) -> int:
        """
        In one dimension there could be several axes to enable main and secondary axis and maybe more in future.

        This method returns the maximum index at which an axis exists for the given dimension. It is allowed that some indexes in between do not have an axis.

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        return self.__component.getMaximumAxisIndexByDimension(dimension)

    def get_view_service_name(self) -> str:
        """
        return a service name from which the view component for this coordinate system can be created
        """
        return self.__component.getViewServiceName()

    def set_axis_by_dimension(self, dimension: int, axis: XAxis, idx: int) -> None:
        """
        The dimension says whether it is a x, y or z axis.

        The index says whether it is a primary or a secondary axis. Use nIndex == 0 for a primary axis.

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        self.__component.setAxisByDimension(dimension, axis, idx)

    # endregion XCoordinateSystem
