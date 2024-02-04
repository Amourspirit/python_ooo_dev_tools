from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.util import XCloneable

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class CloneablePartial:
    """
    Partial Class XCloneable.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XCloneable, interface: UnoInterface | None = XCloneable) -> None:
        """
        Constructor

        Args:
            component (XCloneable): UNO Component that implements ``com.sun.star.util.XCloneable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCloneable``.
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

    # region XCloneable
    def create_clone(self) -> XCloneable:
        """
        Creates a clone of the object.

        Returns:
            XCloneable: The clone.
        """
        return self.__component.createClone()

    # endregion XCloneable
