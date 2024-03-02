from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.table import XAutoFormattable

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class AutoFormattablePartial:
    """
    Partial Class for XAutoFormattable.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XAutoFormattable, interface: UnoInterface | None = XAutoFormattable) -> None:
        """
        Constructor

        Args:
            component (XAutoFormattable): UNO Component that implements ``com.sun.star.table.XAutoFormattable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XAutoFormattable``.
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

    # region XAutoFormattable
    def auto_format(self, name: str) -> None:
        """
        Applies an AutoFormat to the cell range of the current context.

        Args:
            name (str): The name of the AutoFormat to be applied.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
        """
        self.__component.autoFormat(name)

    # endregion XAutoFormattable
