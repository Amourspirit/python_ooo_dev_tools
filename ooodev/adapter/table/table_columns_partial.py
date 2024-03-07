from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from com.sun.star.table import XTableColumns
from ooodev.adapter.container.index_access_partial import IndexAccessPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TableColumnsPartial(IndexAccessPartial):
    """
    Partial Class for XTableColumns.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTableColumns, interface: UnoInterface | None = XTableColumns) -> None:
        """
        Constructor

        Args:
            component (XTableColumns): UNO Component that implements ``com.sun.star.container.XTableColumns`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTableColumns``.
        """
        IndexAccessPartial.__init__(self, component, interface)
        self.__component = component

    # region XTableColumns
    def insert_by_index(self, idx: int, count: int) -> None:
        """
        Inserts a new column at the specified index.

        Args:
            idx (int): The index at which the column will be inserted.
            count (int): The number of columns to insert.
        """
        self.__component.insertByIndex(idx, count)

    def remove_by_index(self, idx: int, count: int) -> None:
        """
        Removes columns from the specified index.

        Args:
            idx (int): The index at which the column will be removed.
            count (int): The number of columns to remove.
        """
        self.__component.removeByIndex(idx, count)

    # endregion XTableColumns
