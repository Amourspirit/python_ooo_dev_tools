from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.form import XConfirmDeleteBroadcaster

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.form import XConfirmDeleteListener
    from ooodev.utils.type_var import UnoInterface


class ConfirmDeleteBroadcasterPartial:
    """
    Partial Class for ``XConfirmDeleteBroadcaster``.
    """

    def __init__(
        self, component: XConfirmDeleteBroadcaster, interface: UnoInterface | None = XConfirmDeleteBroadcaster
    ) -> None:
        """
        Constructor

        Args:
            component (XConfirmDeleteBroadcaster): UNO Component that implements ``com.sun.star.container.XConfirmDeleteBroadcaster``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``None``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XConfirmDeleteBroadcaster
    def add_confirm_delete_listener(self, listener: XConfirmDeleteListener) -> None:
        """
        Remembers the specified listener to receive an event for confirming deletions

        ``XConfirmDeleteListener.confirmDelete()`` is called before a deletion is performed. You may use the event to write your own confirmation messages.
        """
        self.__component.addConfirmDeleteListener(listener)

    def remove_confirm_delete_listener(self, listener: XConfirmDeleteListener) -> None:
        """
        Removes the specified listener.
        """
        self.__component.removeConfirmDeleteListener(listener)

    # endregion XConfirmDeleteBroadcaster
