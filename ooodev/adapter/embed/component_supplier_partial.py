from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.embed import XComponentSupplier

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.type_var import UnoInterface

if TYPE_CHECKING:
    from com.sun.star.util import XCloseable


class ComponentSupplierPartial:
    """
    Partial class for XComponentSupplier.
    """

    def __init__(self, component: XComponentSupplier, interface: UnoInterface | None = XComponentSupplier) -> None:
        """
        Constructor

        Args:
            component (XComponentSupplier): UNO Component that implements ``com.sun.star.embed.XComponentSupplier`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XComponentSupplier``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XComponentSupplier
    def get_component(self) -> XCloseable:
        """
        Allows to get access to a component.

        The component may not support ``com.sun.star.lang.XComponent`` interface.
        """
        return self.__component.getComponent()

    # endregion XComponentSupplier
