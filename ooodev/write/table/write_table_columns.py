from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.adapter.text.table_columns_comp import TableColumnsComp
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial
from ooodev.utils import gen_util as mGenUtil
from ooodev.write.table.write_table_column import WriteTableColumn

if TYPE_CHECKING:
    from com.sun.star.table import XTableColumns
    from ooodev.write.table.write_table import WriteTable


class WriteTableColumns(WriteDocPropPartial, WriteTablePropPartial, TableColumnsComp, LoInstPropsPartial):
    """Represents writer table columns."""

    def __init__(self, owner: WriteTable, component: XTableColumns) -> None:
        """
        Constructor

        Args:
            component (XTableColumns): UNO object that supports ``com.sun.star.text.TableColumns`` service.
        """
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        WriteTablePropPartial.__init__(self, obj=owner)
        LoInstPropsPartial.__init__(self, lo_inst=owner.lo_inst)
        TableColumnsComp.__init__(self, component=component)  # type: ignore
        self._current_index = 0

    def __getitem__(self, key: int) -> WriteTableColumn:
        """
        Gets the form at the specified index.

        Args:
            key (key, int): The index of the column. When getting by index can be a negative value to get from the end.

        Returns:
            WriteTableColumn: The column with the specified index.
        """
        index = self._get_index(key)
        return WriteTableColumn(owner=self, idx=index)

    def __iter__(self) -> WriteTableColumns:
        """
        Iterates over the columns.

        Returns:
            WriteTableColumns: The next column.
        """
        self._current_index = 0
        return self

    def __next__(self) -> WriteTableColumn:
        """
        Gets the next column.

        Returns:
            WriteTableColumn: The next column.
        """
        if self._current_index < len(self):
            self._current_index += 1
            return WriteTableColumn(owner=self, idx=self._current_index - 1)
        raise StopIteration

    def _get_index(self, idx: int, allow_greater: bool = False) -> int:
        """
        Gets the index.

        Args:
            idx (int): Index of sheet. Can be a negative value to index from the end of the list.
            allow_greater (bool, optional): If True and index is greater then the number of
                sheets then the index becomes the next index if sheet were appended. Defaults to False.

        Returns:
            int: Index value.
        """
        count = len(self)
        return mGenUtil.Util.get_index(idx, count, allow_greater)

    def __delitem__(self, key: int) -> None:
        """
        Removes the column at the specified index.

        This is short hand for ``remove_by_index()``

        Args:
            key (key, int): The index of the cell. When getting by index can be a negative value to get from the end.

        See Also:
            - :py:meth:`~ooodev.write.table.write_table_columns.remove_by_index`
        """
        self.remove_by_index(key)

    # region XTableColumns Overrides
    def insert_by_index(self, idx: int, count: int = 1) -> None:
        """
        Inserts a new column at the specified index.

        Args:
            idx (int): The index at which the column will be inserted. A value of ``-1`` will append the column at the end.
            count (int, optional): The number of columns to insert. Defaults to ``1``.
        """
        index = self._get_index(idx, True)
        self.component.insertByIndex(index, count)

    def remove_by_index(self, idx: int, count: int = 1) -> None:
        """
        Removes columns from the specified idx.

        Args:
            idx (int): The index at which the column will be removed. A value of ``-1`` will remove the last column.
            count (int, optional): The number of columns to remove. Defaults to ``1``.
        """
        index = self._get_index(idx)
        self.component.removeByIndex(index, count)

    # endregion XTableColumns Overrides

    def append_columns(self, count: int = 1) -> None:
        """
        Appends columns to the table columns.

        Args:
            count (int, optional): The number of columns to append. Defaults to ``1``.
        """
        self.insert_by_index(-1, count)
