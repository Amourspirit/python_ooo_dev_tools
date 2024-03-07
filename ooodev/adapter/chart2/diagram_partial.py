from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.chart2 import XDiagram

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.chart2 import XColorScheme
    from com.sun.star.beans import XPropertySet
    from com.sun.star.chart2 import XLegend
    from com.sun.star.chart2.data import XDataSource
    from com.sun.star.beans import PropertyValue  # Struct
    from ooodev.utils.type_var import UnoInterface


class DiagramPartial:
    """
    Partial class for XDiagram.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDiagram, interface: UnoInterface | None = XDiagram) -> None:
        """
        Constructor

        Args:
            component (XDiagram): UNO Component that implements ``com.sun.star.chart2.XDiagram`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDiagram``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDiagram
    def get_default_color_scheme(self) -> XColorScheme:
        """
        Gets an ``XColorScheme`` that defines the default colors for data series (or data points) in the diagram.
        """
        return self.__component.getDefaultColorScheme()

    def get_floor(self) -> XPropertySet:
        """
        Gets the property set that determines the visual appearance of the floor if any.

        The floor is the bottom of a 3D diagram. For a 2D diagram None is returned.
        """
        return self.__component.getFloor()

    def get_legend(self) -> XLegend:
        """
        Gets the legend, which may represent data series and other information about a diagram in a separate box.
        """
        return self.__component.getLegend()

    def get_wall(self) -> XPropertySet:
        """
        Gets the property set that determines the visual appearance of the wall.

        The wall is the area behind the union of all coordinate systems used in a diagram.
        """
        return self.__component.getWall()

    def set_default_color_scheme(self, color_scheme: XColorScheme) -> None:
        """
        Sets an ``XColorScheme`` that defines the default colors for data series (or data points) in the diagram.
        """
        return self.__component.setDefaultColorScheme(color_scheme)

    def set_diagram_data(self, data_source: XDataSource, *args: PropertyValue) -> None:
        """
        Sets new data to the diagram.

        For standard parameters that may be used, see the service ``StandardDiagramCreationParameters``.
        """
        self.__component.setDiagramData(data_source, args)

    def set_legend(self, legend: XLegend) -> None:
        """
        Sets a new legend.
        """
        self.__component.setLegend(legend)

    # endregion XDiagram
