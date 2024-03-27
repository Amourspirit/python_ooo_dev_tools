from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.util import XModifyBroadcaster
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.util import XModifyListener


class ModifyBroadcasterPartial:
    """
    Partial Class XModifyBroadcaster.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XModifyBroadcaster, interface: UnoInterface | None = XModifyBroadcaster) -> None:
        """
        Constructor

        Args:
            component (XModifyBroadcaster): UNO Component that implements ``com.sun.star.util.XModifyBroadcaster`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XModifyBroadcaster``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XModifyBroadcaster
    def add_modify_listener(self, listener: XModifyListener) -> None:
        """
        Adds the specified listener to receive events ``modified``.
        """
        self.__component.addModifyListener(listener)

    def remove_modify_listener(self, listener: XModifyListener) -> None:
        """
        Removes the specified listener.
        """
        self.__component.removeModifyListener(listener)

    # endregion XModifyBroadcaster
