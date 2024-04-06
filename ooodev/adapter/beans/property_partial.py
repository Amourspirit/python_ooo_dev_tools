from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.beans import XProperty

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import Property  # struct
    from ooodev.utils.type_var import UnoInterface


class PropertyPartial:
    """
    Partial class for XProperty.
    """

    def __init__(self, component: XProperty, interface: UnoInterface | None = XProperty) -> None:
        """
        Constructor

        Args:
            component (XProperty): UNO Component that implements ``com.sun.star.bean.XProperty`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XProperty``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XProperty

    def get_as_property(self) -> Property:
        """
        Gets the property.
        """
        return self.__component.getAsProperty()

    # endregion XProperty


def get_builder(component: Any) -> Any:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)
    builder.add_import(
        name="ooodev.adapter.beans.property_partial.PropertyPartial",
        uno_name="com.sun.star.beans.XProperty",
        optional=False,
        init_kind=2,
    )
    return builder
