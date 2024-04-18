from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.beans import XMultiPropertySet

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.beans import XPropertiesChangeListener
    from com.sun.star.beans import XPropertySetInfo


class MultiPropertySetPartial:
    """
    Partial class for XMultiPropertySet.
    """

    def __init__(self, component: XMultiPropertySet, interface: UnoInterface | None = XMultiPropertySet) -> None:
        """
        Constructor

        Args:
            component (XMultiPropertySet): UNO Component that implements ``com.sun.star.bean.XMultiPropertySet`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XMultiPropertySet``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XMultiPropertySet

    def add_properties_change_listener(
        self,
        listener: XPropertiesChangeListener,
        *names: str,
    ) -> None:
        """
        adds an XPropertiesChangeListener to the specified property with the specified names.

        The implementation can ignore the names of the properties and fire the event on all properties.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.

        Args:
            listener (XPropertiesChangeListener): The listener to be added.
            names (str): One or more names of the properties to listen to.
        """
        self.__component.addPropertiesChangeListener(names, listener)

    def fire_properties_change_event(self, listener: XPropertiesChangeListener, *names: str) -> None:
        """
        Fires a sequence of PropertyChangeEvents to the specified listener.

        Args:
            listener (XPropertiesChangeListener): The listener to be notified.
            names (str): One or more names of the properties to fire.
        """
        self.__component.firePropertiesChangeEvent(names, listener)

    def get_property_set_info(self) -> XPropertySetInfo:
        """
        Gets the information of the properties which can be set at this object.
        """
        return self.__component.getPropertySetInfo()

    def get_property_values(self, *names: str) -> Tuple[Any, ...]:
        """
        The order of the values in the returned sequence will be the same as the order of the names in the argument.

        Args:
            names (str): One or more names of the properties.
        """
        return self.__component.getPropertyValues(names)

    def remove_properties_change_listener(self, listener: XPropertiesChangeListener) -> None:
        """
        removes an ``XPropertiesChangeListener`` from the listener list.

        It is a ``noop`` if the listener is not registered.

        It is suggested to allow multiple registration of the same listener, thus for each time a listener is added, it has to be removed.
        """
        self.__component.removePropertiesChangeListener(listener)

    def set_property_values(self, names: Tuple[str, ...], values: Tuple[Any, ...]) -> None:
        """
        Sets the values to the properties with the specified names.

        The values of the properties must change before the bound events are fired. The values of the constrained properties should change after the vetoable events are fired and only if no exception occurred. Unknown properties are ignored.

        Similar to ``set_prop_values()`` but with tuples instead of key, value arguments.

        Raises:
            com.sun.star.beans.PropertyVetoException: ``PropertyVetoException``
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.lang.WrappedTargetException: ``WrappedTargetException``
        """
        self.__component.setPropertyValues(names, values)

    def set_prop_values(self, **kwargs: Any) -> None:
        """
        Set the property values.

        Similar to ``set_property_values()`` but with key, value arguments instead of tuples.

        Args:
            kwargs (Any): Property values.
        """
        if not kwargs:
            return
        names, values = zip(*kwargs.items())
        self.__component.setPropertyValues(names, values)

    # endregion XMultiPropertySet


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
    builder.auto_add_interface("com.sun.star.beans.XMultiPropertySet", False)
    return builder
