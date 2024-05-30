from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.document import XActionLockable

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo


if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ActionLockablePartial:
    """
    Partial class for XActionLockable.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XActionLockable, interface: UnoInterface | None = XActionLockable) -> None:
        """
        Constructor

        Args:
            component (XActionLockable): UNO Component that implements ``com.sun.star.document.XActionLockable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XActionLockable``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XActionLockable
    def add_action_lock(self) -> None:
        """
        increments the lock count of the object by one.
        """
        self.__component.addActionLock()

    def is_action_locked(self) -> bool:
        """ """
        return self.__component.isActionLocked()

    def remove_action_lock(self) -> None:
        """
        decrements the lock count of the object by one.
        """
        self.__component.removeActionLock()

    def reset_action_locks(self) -> int:
        """
        resets the locking level.

        This method is used for debugging purposes. The debugging environment of a programming language can reset the locks to allow refreshing of the view if a breakpoint is reached or step execution is used.
        """
        return self.__component.resetActionLocks()

    def set_action_locks(self, lock: int) -> None:
        """
        sets the locking level.

        This method is used for debugging purposes. The programming environment can restore the locking after a break of a debug session.
        """
        self.__component.setActionLocks(lock)

    # endregion XActionLockable
