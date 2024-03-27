from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sdb import XRowSetApproveBroadcaster

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.sdb import XRowSetApproveListener
    from ooodev.utils.type_var import UnoInterface


class RowSetApproveBroadcasterPartial:
    """
    Partial Class for ``XRowSetApproveBroadcaster``.
    """

    def __init__(
        self, component: XRowSetApproveBroadcaster, interface: UnoInterface | None = XRowSetApproveBroadcaster
    ) -> None:
        """
        Constructor

        Args:
            component (XRowSetApproveBroadcaster): UNO Component that implements ``com.sun.star.sdb.XRowSetApproveBroadcaster``.            interface (UnoInterface, optional): The interface to be validated. Defaults to ``None``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XRowSetApproveBroadcaster
    def add_row_set_approve_listener(self, listener: XRowSetApproveListener) -> None:
        """
        Adds the specified listener to receive the events ``approveCursorMove``, ``approveRowChange``, and ``approveRowSetChange``.
        """
        self.__component.addRowSetApproveListener(listener)

    def remove_row_set_approve_listener(self, listener: XRowSetApproveListener) -> None:
        """
        Removes the specified listener.
        """
        self.__component.removeRowSetApproveListener(listener)

    # endregion XRowSetApproveBroadcaster
