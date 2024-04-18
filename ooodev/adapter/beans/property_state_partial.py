from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.beans import XPropertyState
from ooo.dyn.beans.property_state import PropertyState

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class PropertyStatePartial:
    """
    Partial class for XPropertyState.
    """

    def __init__(self, component: XPropertyState, interface: UnoInterface | None = XPropertyState) -> None:
        """
        Constructor

        Args:
            component (XPropertyState): UNO Component that implements ``com.sun.star.bean.XPropertyState`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPropertyState``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPropertyState

    def get_property_default(self, name: str) -> Any:
        """
        If no default exists, is not known or is void, then the return type is void.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        return self.__component.getPropertyDefault(name)

    def get_property_state(self, name: str) -> PropertyState:
        """

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``

        Returns:
            PropertyState: The state of the property.

        Hint:
            - ``PropertyState`` is an enum and can be imported from ``ooo.dyn.beans.property_state``.
        """
        return self.__component.getPropertyState(name)  # type: ignore

    def get_property_states(self, *names: str) -> Tuple[PropertyState, ...]:
        """
        The order of the states is correlating to the order of the given property names.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``

        Returns:
            Tuple[PropertyState, ...]: The states of the properties.

        Hint:
            - ``PropertyState`` is an enum and can be imported from ``ooo.dyn.beans.property_state``.
        """
        return self.__component.getPropertyStates(names)  # type: ignore

    def set_property_to_default(self, name: str) -> None:
        """
        Sets the property to default value.

        The value depends on the implementation of this interface. If it is a bound property, you must change the value before the change events are fired. If it is a constrained property, you must fire the vetoable event before you change the property value.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
        """
        self.__component.setPropertyToDefault(name)

    # endregion XPropertyState


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
        name="ooodev.adapter.beans.property_state_partial.PropertyStatePartial",
        uno_name="com.sun.star.beans.XPropertyState",
        optional=False,
        init_kind=2,
    )
    return builder
