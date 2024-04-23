from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.container import XSet
from ooodev.adapter.container.enumeration_access_partial import EnumerationAccessPartial

if TYPE_CHECKING:
    from ooodev.utils.builder.default_builder import DefaultBuilder
    from ooodev.utils.type_var import UnoInterface


class SetPartial(EnumerationAccessPartial):
    """
    Partial Class for XSet.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XSet, interface: UnoInterface | None = XSet) -> None:
        """
        Constructor

        Args:
            component (XSet): UNO Component that implements ``com.sun.star.container.XSet``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSet``.
        """
        EnumerationAccessPartial.__init__(self, component, interface)
        self.__component = component

    # region XSet
    def has(self, element: Any) -> bool:
        """
        Gets if element is in this container.
        """
        return self.__component.has(element)

    def insert(self, element: Any) -> None:
        """
        Inserts the given element into this container.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.ElementExistException: ``ElementExistException``
        """
        self.__component.insert(element)

    def remove(self, element: Any) -> None:
        """
        removes the given element from this container.

        Raises:
            com.sun.star.lang.IllegalArgumentException: ``IllegalArgumentException``
            com.sun.star.container.NoSuchElementException: ``NoSuchElementException``
        """
        self.__component.remove(element)

    # endregion XSet


def get_builder(component: Any) -> DefaultBuilder:
    """
    Get the builder for the component.

    Args:
        component (Any): The component.

    Returns:
        DefaultBuilder: Builder instance.
    """
    # pylint: disable=import-outside-toplevel
    # pylint: disable=redefined-outer-name
    from ooodev.utils.builder.default_builder import DefaultBuilder

    builder = DefaultBuilder(component)
    builder.auto_add_interface("com.sun.star.container.XSet")
    builder.set_omit(
        "com.sun.star.container.XEnumeration",
        "com.sun.star.container.XElementAccess",
    )

    return builder
