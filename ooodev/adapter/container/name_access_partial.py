from __future__ import annotations
from typing import Any, Generic, TypeVar
import uno

from com.sun.star.container import XNameAccess

from ooodev.utils.type_var import UnoInterface
from ooodev.adapter.container import element_access_partial
from ooodev.utils.builder.default_builder import DefaultBuilder

T = TypeVar("T")


class NameAccessPartial(Generic[T], element_access_partial.ElementAccessPartial):
    """
    Partial Class for XNameAccess.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNameAccess, interface: UnoInterface | None = XNameAccess) -> None:
        """
        Constructor

        Args:
            component (XNameAccess): UNO Component that implements ``com.sun.star.container.XNameAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XNameAccess``.
        """
        element_access_partial.ElementAccessPartial.__init__(self, component, interface)
        self.__component = component

    # region Methods
    def get_by_name(self, name: str) -> T:
        """
        Gets the element with the specified name.

        Args:
            name (str): The name of the element.

        Returns:
            Any: The element with the specified name.
        """
        return self.__component.getByName(name)

    def get_element_names(self) -> tuple[str, ...]:
        """
        Gets the names of all elements contained in the container.

        Returns:
            tuple[str, ...]: The names of all elements.
        """
        return tuple(self.__component.getElementNames())

    def has_by_name(self, name: str) -> bool:
        """
        Checks if the container has an element with the specified name.

        Args:
            name (str): The name of the element.

        Returns:
            bool: ``True`` if the container has an element with the specified name, otherwise ``False``.
        """
        return self.__component.hasByName(name)

    # endregion Methods

    def __contains__(self, name: str) -> bool:
        """
        Checks if the container has an element with the specified name.

        Args:
            name (str): The name of the element.

        Returns:
            bool: ``True`` if the container has an element with the specified name, otherwise ``False``.
        """
        return self.has_by_name(name)


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """

    builder = DefaultBuilder(component)
    builder.set_omit("com.sun.star.container.XElementAccess")
    builder.auto_add_interface("com.sun.star.container.XNameAccess")
    return builder
