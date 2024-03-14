from __future__ import annotations
from typing import TYPE_CHECKING, Generic, TypeVar

from com.sun.star.container import XNameContainer
from ooodev.adapter.container.name_access_partial import NameAccessPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface

T = TypeVar("T")


class NameContainerPartial(NameAccessPartial[T], Generic[T]):
    """
    Partial Class for XNameContainer.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XNameContainer, interface: UnoInterface | None = XNameContainer) -> None:
        """
        Constructor

        Args:
            component (XNameContainer): UNO Component that implements ``com.sun.star.container.XNameContainer``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XNameContainer``.
        """
        NameAccessPartial.__init__(self, component, interface)
        self.__component = component

    # region XNameContainer
    def insert_by_name(self, name: str, element: T) -> None:
        """Inserts the element with the specified name.

        Args:
            name (str): The name of the element to be inserted.
            element (T): The new element.
        """
        self.__component.insertByName(name, element)

    def remove_by_name(self, name: str) -> None:
        """Removes the element with the specified name.

        Args:
            name (str): The name of the element to be removed.
        """
        self.__component.removeByName(name)

    # endregion XNameContainer
