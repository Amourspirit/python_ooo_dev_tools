from __future__ import annotations
from typing import Any
import uno

from com.sun.star.container import XEnumeration

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo
from ooodev.utils.type_var import UnoInterface


class EnumerationPartial:
    """
    Partial class for XEnumeration.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XEnumeration, interface: UnoInterface | None = XEnumeration) -> None:
        """
        Constructor

        Args:
            component (XEnumeration): UNO Component that implements ``com.sun.star.container.XEnumeration`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XEnumeration``.
        """

        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
            interface (UnoInterface): The interface to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XEnumeration
    def has_more_elements(self) -> bool:
        """
        tests whether this enumeration contains more elements.
        """
        return self.__component.hasMoreElements()

    def next_element(self) -> Any:
        """
        Gets the next element of this enumeration.

        Returns:
            Any: the next element of this enumeration.
        """
        return self.__component.nextElement()

    # endregion XEnumeration
