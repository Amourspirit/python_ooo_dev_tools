from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.beans import XPropertySetInfo

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.beans import Property  # struct
    from ooodev.utils.type_var import UnoInterface


class PropertySetInfoPartial:
    """
    Partial class for XPropertySetInfo.
    """

    def __init__(self, component: XPropertySetInfo, interface: UnoInterface | None = XPropertySetInfo) -> None:
        """
        Constructor

        Args:
            component (XPropertySetInfo): UNO Component that implements ``com.sun.star.bean.XPropertySetInfo`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPropertySetInfo``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPropertySetInfo

    def get_properties(self) -> Tuple[Property, ...]:
        """
        Gets all properties of the object.
        """
        return self.__component.getProperties()

    def get_property_by_name(self, name: str) -> Property:
        """
        Gets the property with the specified name.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
        """
        return self.__component.getPropertyByName(name)

    def has_property_by_name(self, name: str) -> bool:
        """
        Gets if the object has a property with the specified name.
        """
        return self.__component.hasPropertyByName(name)

    # endregion XPropertySetInfo


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
        name="ooodev.adapter.beans.property_set_info_partial.PropertySetInfoPartial",
        uno_name="com.sun.star.beans.XPropertySetInfo",
        optional=False,
        init_kind=2,
    )
    return builder
