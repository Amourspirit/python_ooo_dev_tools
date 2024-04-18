from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.beans import XMultiPropertyStates
from ooo.dyn.beans.property_state import PropertyState

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class MultiPropertyStatesPartial:
    """
    Partial class for XMultiPropertyStates.
    """

    def __init__(self, component: XMultiPropertyStates, interface: UnoInterface | None = XMultiPropertyStates) -> None:
        """
        Constructor

        Args:
            component (XMultiPropertyStates): UNO Component that implements ``com.sun.star.bean.XMultiPropertyStates`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XMultiPropertyStates``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XMultiPropertyStates

    def get_property_defaults(self, *names: str) -> Tuple[Any, ...]:
        """
        If no default exists, is not known, or is void, then the return type at the corresponding position in the sequence returned is void.

        Args:
            names (str): One or more names of the properties.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``

        Returns:
            Tuple[Any, ...]: The default values of the properties.
        """
        return self.__component.getPropertyDefaults(names)

    def get_property_states(self, *names: str) -> Tuple[PropertyState, ...]:
        """
        The order of the states is correlating to the order of the given property names.

        Args:
            names (str): One or more names of the properties.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``

        Returns:
            Tuple[PropertyState, ...]: The states of the properties.
        """
        return self.__component.getPropertyStates(names)

    def set_all_properties_to_default(self) -> None:
        """
        Sets all properties to their default values.

        Each value depends on the implementation of this interface. If it is a bound property, you must change the value before the change events are fired. If it is a constrained property, you must fire the vetoable event before you change the property value.
        """
        self.__component.setAllPropertiesToDefault()

    def set_properties_to_default(self, *names: str) -> None:
        """
        Sets the specified properties to their default values.

        Each value depends on the implementation of this interface. If it is a bound property, you must change the value before the change events are fired. If it is a constrained property, you must fire the vetoable event before you change the property value.

        Args:
            names (str): One or more names of the properties.

        Raises:
            com.sun.star.beans.UnknownPropertyException: ``UnknownPropertyException``
        """
        self.__component.setPropertiesToDefault(names)

    # endregion XMultiPropertyStates


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
        name="ooodev.adapter.beans.multi_property_states_partial.MultiPropertyStatesPartial",
        uno_name="com.sun.star.beans.XMultiPropertyStates",
        optional=False,
        init_kind=2,
    )
    return builder
