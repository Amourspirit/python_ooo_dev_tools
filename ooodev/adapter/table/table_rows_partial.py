from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.table import XTableRows
from ooodev.adapter.container.index_access_partial import IndexAccessPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TableRowsPartial(IndexAccessPartial):
    """
    Partial Class for XTableRows.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTableRows, interface: UnoInterface | None = XTableRows) -> None:
        """
        Constructor

        Args:
            component (XTableRows): UNO Component that implements ``com.sun.star.container.XTableRows`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTableRows``.
        """
        IndexAccessPartial.__init__(self, component, interface)
        self.__component = component

    # region XTableRows
    def insert_by_index(self, index: int, count: int) -> None:
        """
        Inserts a new column at the specified index.

        Args:
            index (int): The index at which the column will be inserted.
            count (int): The number of columns to insert.
        """
        self.__component.insertByIndex(index, count)

    def remove_by_index(self, index: int, count: int) -> None:
        """
        Removes columns from the specified index.

        Args:
            index (int): The index at which the column will be removed.
            count (int): The number of columns to remove.
        """
        self.__component.removeByIndex(index, count)

    # endregion XTableRows
