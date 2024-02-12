from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.graphic import XGraphic

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class GraphicPartial:
    """
    Partial class for XGraphic.
    """

    def __init__(self, component: XGraphic, interface: UnoInterface | None = XGraphic) -> None:
        """
        Constructor

        Args:
            component (XGraphic): UNO Component that implements ``com.sun.star.frame.XGraphic`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XGraphic``.
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

    # region XGraphic
    def get_type(self) -> int:
        """
        Get the type of the contained graphic.
        """
        return self.__component.getType()

    # endregion XGraphic
