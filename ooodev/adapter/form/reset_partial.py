from __future__ import annotations
from typing import TYPE_CHECKING

from com.sun.star.form import XReset

if TYPE_CHECKING:
    from com.sun.star.form import XResetListener
    from ooodev.utils.type_var import UnoInterface


class ResetPartial:
    """
    Partial Class for XReset.

    This interface does not really provide an own functionality, it is only for easier runtime identification of form components.
    """

    def __init__(self, component: XReset, interface: UnoInterface | None = XReset) -> None:
        """
        Constructor

        Args:
            component (XReset): UNO Component that implements ``com.sun.star.container.XReset``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XReset``.
        """
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
