from __future__ import annotations
from typing import Any
import uno

from com.sun.star.container import XIndexReplace

from ooodev.utils.type_var import UnoInterface
from ooodev.adapter.container.index_access_partial import IndexAccessPartial


class IndexReplacePartial(IndexAccessPartial):
    """
    Partial Class for XIndexReplace.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XIndexReplace, interface: UnoInterface | None = XIndexReplace) -> None:
        """
        Constructor

        Args:
            component (XIndexReplace): UNO Component that implements ``com.sun.star.container.XIndexReplace`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XNameAccess``.
        """
        IndexAccessPartial.__init__(self, component, interface)
        self.__component = component

    # region XIndexReplace
    def replace_by_index(self, index: int, element: Any) -> None:
        """
        Replaces the element at the specified index with the given element.

        Args:
            index (int): The index of the element that is to be replaced.
            element (Any): The replacement element.

        Returns:
            None:
        """
        self.__component.replaceByIndex(index, element)

    # endregion XIndexReplace
