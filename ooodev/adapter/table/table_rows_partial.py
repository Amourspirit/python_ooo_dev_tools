from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno

from com.sun.star.table import XTableRows
from ooodev.adapter.container.index_access_partial import IndexAccessPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface

T = TypeVar("T")


class TableRowsPartial(IndexAccessPartial[T], Generic[T]):
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
    def insert_by_index(self, idx: int, count: int = 1) -> None:
        """
        Inserts a new column at the specified index.

        Args:
            idx (int): The index at which the column will be inserted.
            count (int, optional): The number of columns to insert. Defaults to ``1``.
        """
        self.__component.insertByIndex(idx, count)

    def remove_by_index(self, idx: int, count: int = 1) -> None:
        """
        Removes columns from the specified index.

        Args:
            idx (int): The index at which the column will be removed.
            count (int, optional): The number of columns to remove. Default is ``1``.
        """
        self.__component.removeByIndex(idx, count)

    # endregion XTableRows
