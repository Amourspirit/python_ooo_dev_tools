from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.beans import XPropertyWithState
from ooo.dyn.beans.property_state import PropertyState

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.uno import XInterface
    from ooodev.utils.type_var import UnoInterface


class PropertyWithStatePartial:
    """
    Partial class for XPropertyWithState.
    """

    def __init__(self, component: XPropertyWithState, interface: UnoInterface | None = XPropertyWithState) -> None:
        """
        Constructor

        Args:
            component (XPropertyWithState): UNO Component that implements ``com.sun.star.bean.XPropertyWithState`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XPropertyWithState``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XPropertyWithState
    def get_default_as_property(self) -> XInterface:
        """
        If no default exists, is not known or is void, then the return value is NULL.

        Raises:
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        return self.__component.getDefaultAsProperty()

    def get_state_as_property(self) -> PropertyState:
        """
        Gets the state of the property.

        Returns:
            PropertyState: The state of the property.

        Hint:
            - ``PropertyState`` is an enum and can be imported from ``ooo.dyn.beans.property_state``.
        """
        return self.__component.getStateAsProperty()  # type: ignore

    def set_to_default_as_property(self) -> None:
        """
        Sets this to its default value.

        The value depends on the implementation of this interface. If this is a bound property, the value changes before the change events are fired. If this is a constrained property, the vetoable event is fired before the property value changes.

        Raises:
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self.__component.setToDefaultAsProperty()

    # endregion XPropertyWithState


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
        name="ooodev.adapter.beans.property_with_state_partial.PropertyWithStatePartial",
        uno_name="com.sun.star.beans.XPropertyWithState",
        optional=False,
        init_kind=2,
    )
    return builder
