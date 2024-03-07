from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.beans import XPropertyBag

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.adapter.beans.property_container_partial import PropertyContainerPartial
from ooodev.adapter.beans.property_access_partial import PropertyAccessPartial

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

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        PropertySetPartial.__init__(self, component, interface)
        PropertyContainerPartial.__init__(self, component, interface)
        PropertyAccessPartial.__init__(self, component, interface)
