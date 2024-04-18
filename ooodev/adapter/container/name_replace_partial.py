from __future__ import annotations
from typing import Any, TYPE_CHECKING, Generic, TypeVar

from com.sun.star.container import XNameReplace
from ooodev.adapter.container.name_access_partial import NameAccessPartial
from ooodev.utils.builder.default_builder import DefaultBuilder

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface

T = TypeVar("T")


class NameReplacePartial(NameAccessPartial[T], Generic[T]):
    """
    Partial Class for XNameReplace.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNameReplace, interface: UnoInterface | None = XNameReplace) -> None:
        """
        Constructor

        Args:
            component (XNameReplace): UNO Component that implements ``com.sun.star.container.XNameReplace``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XNameReplace``.
        """
        NameAccessPartial.__init__(self, component, interface)
        self.__component = component

    # region XNameReplace
    def replace_by_name(self, name: str, element: T) -> None:
        """Replaces the element with the specified name.

        Args:
            name (str): The name of the element to be replaced.
            element (T): The new element.
        """
        self.__component.replaceByName(name, element)

    # endregion XNameReplace


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
    builder.set_omit("com.sun.star.container.XNameAccess")
    builder.auto_add_interface("com.sun.star.container.XNameReplace")
    return builder
