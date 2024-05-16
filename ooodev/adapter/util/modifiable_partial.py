from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.util import XModifiable
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.util.modify_broadcaster_partial import ModifyBroadcasterPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ModifiablePartial(ModifyBroadcasterPartial):
    """
    Partial Class XModifiable.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XModifiable, interface: UnoInterface | None = XModifiable) -> None:
        """
        Constructor

        Args:
            component (XModifiable): UNO Component that implements ``com.sun.star.util.XModifiable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XModifiable``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        ModifyBroadcasterPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XModifiable
    def is_modified(self) -> bool:
        """
        The modification is always in relation to a certain state (i.e., the initial, loaded, or last stored version).
        """
        return self.__component.isModified()

    def set_modified(self, modified: bool) -> None:
        """
        sets the status of the modified-flag from outside of the object.

        Raises:
            com.sun.star.beans.PropertyVetoException: ``PropertyVetoException``
        """
        self.__component.setModified(modified)

    # endregion XModifiable
