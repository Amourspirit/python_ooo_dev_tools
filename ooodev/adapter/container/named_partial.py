from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

if TYPE_CHECKING:
    from com.sun.star.container import XNamed

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo


class NamedPartial:
    """
    Class for managing XNamed.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNamed) -> None:
        """
        Constructor

        Args:
            component (XNamed): UNO Component that implements ``com.sun.star.container.XNamed`` interface.
        """

        if not mLo.Lo.is_uno_interfaces(component, XNamed):
            raise mEx.MissingInterfaceError("XNamed")
        self.__component = component

    # region Methods
    def get_name(self) -> str:
        """Returns the name of the object."""
        return self.__component.getName()

    def set_name(self, name: str) -> None:
        """Sets the name of the object."""
        self.__component.setName(name)

    # endregion Methods
