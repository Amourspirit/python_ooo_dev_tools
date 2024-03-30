from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.form import XBoundControl

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class BoundControlPartial:
    """
    Partial Class for XBoundControl.
    """

    def __init__(self, component: XBoundControl, interface: UnoInterface | None = XBoundControl) -> None:
        """
        Constructor

        Args:
            component (XBoundControl): UNO Component that implements ``com.sun.star.container.XBoundControl``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``None``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XBoundControl
    def get_lock(self) -> bool:
        """
        Gets whether the input is currently locked or not.
        """
        return self.__component.getLock()

    def set_lock(self, lock: bool) -> None:
        """
        is used for altering the current lock state of the component.
        """
        return self.__component.setLock(lock)

    # endregion XBoundControl
