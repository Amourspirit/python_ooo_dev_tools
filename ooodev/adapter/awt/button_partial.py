from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XButton

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XActionListener
    from ooodev.utils.type_var import UnoInterface


class ButtonPartial:
    """
    Partial class for XButton.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XButton, interface: UnoInterface | None = XButton) -> None:
        """
        Constructor

        Args:
            component (XButton): UNO Component that implements ``com.sun.star.awt.XButton`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XButton``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XButton
    def add_action_listener(self, listener: XActionListener) -> None:
        """
        Registers an event handler for button action events.
        """
        self.__component.addActionListener(listener)

    def remove_action_listener(self, listener: XActionListener) -> None:
        """
        Un-registers an event handler for button action events.
        """
        self.__component.removeActionListener(listener)

    def set_action_command(self, cmd: str) -> None:
        """
        Sets a command string for pushing the button.
        """
        self.__component.setActionCommand(cmd)

    def set_label(self, label: str) -> None:
        """
        Sets the label of the button.
        """
        self.__component.setLabel(label)

    # endregion XButton
