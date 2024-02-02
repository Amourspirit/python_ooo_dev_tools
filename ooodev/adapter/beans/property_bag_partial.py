from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.beans import XPropertyBag

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from .property_set_partial import PropertySetPartial
from .property_container_partial import PropertyContainerPartial
from .property_access_partial import PropertyAccessPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class PropertyBagPartial(PropertySetPartial, PropertyContainerPartial, PropertyAccessPartial):
    """
    Partial class for XPropertyBag.
    """

    def __init__(self, component: XPropertyBag, interface: UnoInterface | None = XPropertyBag) -> None:
        """
        Constructor

        Args:
            component (XPropertyBag): UNO Component that implements ``com.sun.star.container.XPropertyBag`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPropertyBag``.
        """

        self.__interface = interface
        self.__validate(component)
        PropertySetPartial.__init__(self, component, interface)
        PropertyContainerPartial.__init__(self, component, interface)
        PropertyAccessPartial.__init__(self, component, interface)

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)
