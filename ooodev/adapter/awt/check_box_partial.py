from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XCheckBox

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.kind.tri_state_kind import TriStateKind

if TYPE_CHECKING:
    from com.sun.star.awt import XItemListener
    from ooodev.utils.type_var import UnoInterface


class CheckBoxPartial:
    """
    Partial class for XCheckBox.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XCheckBox, interface: UnoInterface | None = XCheckBox) -> None:
        """
        Constructor

        Args:
            component (XCheckBox): UNO Component that implements ``com.sun.star.awt.XCheckBox`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCheckBox``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XCheckBox
    def add_item_listener(self, listener: XItemListener) -> None:
        """
        Registers a listener for item events.
        """
        self.__component.addItemListener(listener)

    def enable_tri_state(self, enable: bool) -> None:
        """
        Enables or disables the tri state mode.
        """
        self.__component.enableTriState(enable)

    def get_state(self) -> TriStateKind:
        """
        Gets the state of the check box.

        Returns:
            TriStateKind: The state of the check box.

        Hint:
            - ``TriStateKind`` is an enum and can be imported ``ooodev.utils.kind.tri_state_kind``.
        """
        return TriStateKind(self.__component.getState())

    def remove_item_listener(self, listener: XItemListener) -> None:
        """
        Un-registers a listener for item events.
        """
        self.__component.removeItemListener(listener)

    def set_label(self, label: str) -> None:
        """

        Sets the label of the check box.
        """
        self.__component.setLabel(label)

    def set_state(self, state: int | TriStateKind) -> None:
        """
        Sets the state of the check box.
        """
        self.__component.setState(int(state))

    # endregion XCheckBox
