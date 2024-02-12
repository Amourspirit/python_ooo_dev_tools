from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.chart2 import XCoordinateSystemContainer

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.chart2 import XCoordinateSystem
    from ooodev.utils.type_var import UnoInterface


class CoordinateSystemContainerPartial:
    """
    Partial class for XCoordinateSystemContainer.
    """

    def __init__(
        self, component: XCoordinateSystemContainer, interface: UnoInterface | None = XCoordinateSystemContainer
    ) -> None:
        """
        Constructor

        Args:
            component (XCoordinateSystemContainer): UNO Component that implements ``com.sun.star.chart2.XCoordinateSystemContainer`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCoordinateSystemContainer``.
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

    # region XCoordinateSystemContainer
    def add_coordinate_system(self, coord_sys: XCoordinateSystem) -> None:
        """
        Add a coordinate system to the coordinate system container

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.addCoordinateSystem(coord_sys)

    def get_coordinate_systems(self) -> Tuple[XCoordinateSystem, ...]:
        """
        Gets all coordinate systems
        """
        return self.__component.getCoordinateSystems()

    def remove_coordinate_system(self, coord_sys: XCoordinateSystem) -> None:
        """
        Removes one coordinate system from the coordinate system container.

        Raises:
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        self.__component.removeCoordinateSystem(coord_sys)

    def set_coordinate_systems(self, *systems: XCoordinateSystem) -> None:
        """
        Set all coordinate systems.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.setCoordinateSystems(systems)

    # endregion XCoordinateSystemContainer
