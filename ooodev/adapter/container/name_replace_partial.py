from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.container import XNameReplace
from ooodev.adapter.container.name_access_partial import NameAccessPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class NameReplacePartial(NameAccessPartial):
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
    def replace_by_name(self, name: str, element: Any) -> None:
        """Replaces the element with the specified name.

        Args:
            name (str): The name of the element to be replaced.
            element (object): The new element.
        """
        self.__component.replaceByName(name, element)

    # endregion XNameReplace
