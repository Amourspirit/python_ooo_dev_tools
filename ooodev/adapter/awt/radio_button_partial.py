from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XRadioButton

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XItemListener
    from ooodev.utils.type_var import UnoInterface


class RadioButtonPartial:
    """
    Partial class for XRadioButton.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XRadioButton, interface: UnoInterface | None = XRadioButton) -> None:
        """
        Constructor

        Args:
            component (XRadioButton): UNO Component that implements ``com.sun.star.awt.XRadioButton`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XRadioButton``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XRadioButton
    def add_item_listener(self, listener: XItemListener) -> None:
        """
        Registers a listener for item events.
        """
        self.__component.addItemListener(listener)

    def get_state(self) -> bool:
        """
        Returns ``True`` if the button is checked, ``False`` otherwise.
        """
        return self.__component.getState()

    def remove_item_listener(self, listener: XItemListener) -> None:
        """
        Un-registers a listener for item events.
        """
        self.__component.removeItemListener(listener)

    def set_label(self, label: str) -> None:
        """

        Sets the label of the radio button.
        """
        self.__component.setLabel(label)

    def set_state(self, state: bool) -> None:
        """
        Sets the state of the radio button.
        """
        self.__component.setState(state)

    # endregion XRadioButton
