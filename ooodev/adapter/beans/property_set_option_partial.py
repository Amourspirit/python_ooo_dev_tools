from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.beans import XPropertySetOption

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import Property  # struct
    from ooodev.utils.type_var import UnoInterface


class PropertySetOptionPartial:
    """
    Partial class for XPropertySetOption.
    """

    def __init__(self, component: XPropertySetOption, interface: UnoInterface | None = XPropertySetOption) -> None:
        """
        Constructor

        Args:
            component (XPropertySetOption): UNO Component that implements ``com.sun.star.bean.XPropertySetOption`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPropertySetOption``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPropertySetOption
    def enable_change_listener_notification(self, enable: bool) -> None:
        """
        Turn on or off notifying change listeners on property value change.

        This option is turned on by default.
        """
        self.__component.enableChangeListenerNotification(enable)

    # endregion XPropertySetOption
