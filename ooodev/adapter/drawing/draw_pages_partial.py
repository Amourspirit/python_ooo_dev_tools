from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.drawing import XDrawPages

from ooodev.adapter.container.index_access_partial import IndexAccessPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage
    from ooodev.utils.type_var import UnoInterface


class DrawPagesPartial(IndexAccessPartial["XDrawPage"]):
    """
    Partial Class for XDrawPages.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XDrawPages, interface: UnoInterface | None = XDrawPages) -> None:
        """
        Constructor

        Args:
            component (XDrawPages): UNO Component that implements ``com.sun.star.drawing.XDrawPages`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDrawPages``.
        """
        IndexAccessPartial.__init__(self, component, interface)
        self.__component = component

    # region XDrawPages
    def insert_new_by_index(self, idx: int) -> XDrawPage:
        """
        Creates and inserts a new DrawPage or MasterPage into this container.

        Args:
            idx (int): The index at which the new page will be inserted.

        Returns:
            Any: The new page.
        """
        return self.__component.insertNewByIndex(idx)

    def remove(self, page: Any) -> None:
        """
        Removes a DrawPage or MasterPage from this container.

        Args:
            page (Any): The page to be removed.
        """
        self.__component.remove(page)

    # endregion XDrawPages
