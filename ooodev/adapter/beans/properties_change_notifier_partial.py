from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.beans import XPropertiesChangeNotifier

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.beans import XPropertiesChangeListener


class PropertiesChangeNotifierPartial:
    """
    Partial class for XPropertiesChangeNotifier.
    """

    def __init__(
        self, component: XPropertiesChangeNotifier, interface: UnoInterface | None = XPropertiesChangeNotifier
    ) -> None:
        """
        Constructor

        Args:
            component (XPropertiesChangeNotifier): UNO Component that implements ``com.sun.star.bean.XPropertiesChangeNotifier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPropertiesChangeNotifier``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPropertiesChangeNotifier
    def add_properties_change_listener(
        self,
        listener: XPropertiesChangeListener,
        *names: str,
    ) -> None:
        """
        adds an XPropertiesChangeListener to the specified properties with the specified names.
        """
        self.__component.addPropertiesChangeListener(names, listener)

    def remove_properties_change_listener(self, listener: XPropertiesChangeListener, *names: str) -> None:
        """
        Removes an XPropertiesChangeListener from the listener list.
        """
        self.__component.removePropertiesChangeListener(names, listener)

    # endregion XPropertiesChangeNotifier
