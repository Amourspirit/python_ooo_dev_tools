from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.form import XReset

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.builder.default_builder import DefaultBuilder

if TYPE_CHECKING:
    from com.sun.star.form import XResetListener
    from ooodev.utils.type_var import UnoInterface


class ResetPartial:
    """
    Partial Class for XReset.
    """

    def __init__(self, component: XReset, interface: UnoInterface | None = XReset) -> None:
        """
        Constructor

        Args:
            component (XReset): UNO Component that implements ``com.sun.star.container.XReset``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``None``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XReset
    def add_reset_listener(self, listener: XResetListener) -> None:
        """
        Adds a listener to be notified when the form is reset.

        Args:
            listener (XResetListener): Listener to be added.
        """
        self.__component.addResetListener(listener)

    def remove_reset_listener(self, listener: XResetListener) -> None:
        """
        Removes a listener from the list of listeners that are notified when the form is reset.

        Args:
            listener (XResetListener): Listener to be removed.
        """
        self.__component.removeResetListener(listener)

    def reset(self) -> None:
        """
        Resets a component to some default value.
        """
        self.__component.reset()

    # endregion XReset


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    builder = DefaultBuilder(component)

    builder.auto_add_interface("com.sun.star.form.XReset")
    return builder
