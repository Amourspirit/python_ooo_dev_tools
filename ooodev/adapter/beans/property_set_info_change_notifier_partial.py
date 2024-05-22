from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.beans import XPropertySetInfoChangeNotifier

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.beans import XPropertySetInfoChangeListener


class PropertySetInfoChangeNotifierPartial:
    """
    Partial class for XPropertySetInfoChangeNotifier.
    """

    def __init__(
        self,
        component: XPropertySetInfoChangeNotifier,
        interface: UnoInterface | None = XPropertySetInfoChangeNotifier,
    ) -> None:
        """
        Constructor

        Args:
            component (XPropertySetInfoChangeNotifier): UNO Component that implements ``com.sun.star.bean.XPropertySetInfoChangeNotifier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPropertySetInfoChangeNotifier``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPropertySetInfoChangeNotifier
    def add_property_set_info_change_listener(self, listener: XPropertySetInfoChangeListener) -> None:
        """
        Registers a listener for ``PropertySetInfoChangeEvents``.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.
        """
        self.__component.addPropertySetInfoChangeListener(listener)

    def remove_property_set_info_change_listener(self, listener: XPropertySetInfoChangeListener) -> None:
        """
        Removes a listener for ``PropertySetInfoChangeEvents``.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.
        """
        self.__component.removePropertySetInfoChangeListener(listener)

    # endregion XPropertySetInfoChangeNotifier
