from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.ucb import XCommandInfoChangeNotifier

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.ucb import XCommandInfoChangeListener
    from ooodev.utils.type_var import UnoInterface


class CommandInfoChangeNotifierPartial:
    """
    Partial Class XCommandInfoChangeNotifier.
    """

    # pylint: disable=unused-argument

    def __init__(
        self, component: XCommandInfoChangeNotifier, interface: UnoInterface | None = XCommandInfoChangeNotifier
    ) -> None:
        """
        Constructor

        Args:
            component (XCommandInfoChangeNotifier): UNO Component that implements ``com.sun.star.ucb.XCommandInfoChangeNotifier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCommandInfoChangeNotifier``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XCommandInfoChangeNotifier
    def add_command_info_change_listener(self, listener: XCommandInfoChangeListener) -> None:
        """
        registers a listener for CommandInfoChangeEvents.
        """
        self.__component.addCommandInfoChangeListener(listener)

    def remove_command_info_change_listener(self, listener: XCommandInfoChangeListener) -> None:
        """
        removes a listener for CommandInfoChangeEvents.
        """
        self.__component.removeCommandInfoChangeListener(listener)

    # endregion XCommandInfoChangeNotifier
