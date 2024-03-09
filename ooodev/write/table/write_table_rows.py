from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.adapter.text.table_rows_comp import TableRowsComp
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.utils import gen_util as mGenUtil
from ooodev.write.table.write_table_row import WriteTableRow
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial


if TYPE_CHECKING:
    from com.sun.star.table import XTableRows
    from ooodev.write.table.write_table import WriteTable


class WriteTableRows(WriteDocPropPartial, WriteTablePropPartial, TableRowsComp, LoInstPropsPartial):
    """Represents writer table rows."""

    def __init__(self, owner: WriteTable[Any], component: XTableRows) -> None:
        """
        Constructor

        Args:
            component (XTableRows): UNO object that supports ``com.sun.star.table.TableRows`` interface.
        """
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        WriteTablePropPartial.__init__(self, obj=owner)
        LoInstPropsPartial.__init__(self, lo_inst=owner.lo_inst)
        TableRowsComp.__init__(self, component=component)  # type: ignore

    def __next__(self) -> WriteTableRow:
        """
        Gets the next form.

        Returns:
            WriteForm: The next form.
        """
        return WriteTableRow(owner=self, component=super().__next__())

    def __getitem__(self, key: int) -> WriteTableRow:
        """
        Gets the form at the specified index.

        This is short hand for ``get_by_index()``

        Args:
            key (key, int): The index of the row. When getting by index can be a negative value to get from the end.

        Returns:
            WriteTableRow: The row with the specified index.

        See Also:
            - :py:meth:`~ooodev.write.table.write_table_rows.get_by_index`
        """
        return self.get_by_index(key)

    def __delitem__(self, key: int) -> None:
        """
        Removes the row at the specified index.

        This is short hand for ``remove_by_index()``

        Args:
            key (key, int): The index of the cell. When getting by index can be a negative value to get from the end.

        See Also:
            - :py:meth:`~ooodev.write.table.write_table_rows.remove_by_index`
        """
        self.remove_by_index(key)

    # region IndexAccessPartial overrides
    def get_by_index(self, idx: int) -> WriteTableRow:
        """
        Gets the row at the specified index.

        Args:
            idx (int): Index of row.

        Returns:
            Any: Row at the specified index.
        """
        index = self._get_index(idx)
        return WriteTableRow(owner=self, component=self.component.getByIndex(index), idx=index)

    # endregion IndexAccessPartial overrides

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

    # region TableRowsPartial Overrides

    def insert_by_index(self, idx: int, count: int = 1) -> None:
        """
        Inserts rows at the specified index.

        Args:
            idx (int): Index to insert the rows. Idx can be a negative value insert from the end.
                A value of ``-1`` will insert at the end.
            count (int, optional): Number of rows to insert. Defaults to ``1``.

        Returns:
            list[WriteTableRow]: List of inserted rows.
        """
        index = self._get_index(idx, allow_greater=True)
        self.component.insertByIndex(index, count)

    def remove_by_index(self, idx: int, count: int = 1) -> None:
        """
        Removes columns from the specified index.

        Args:
            idx (int): The index at which the column will be removed. Idx can be a negative value insert from the end.
                A value of ``-1`` will remove from the end.
            count (int, optional): The number of columns to remove. Default is ``1``.
        """
        index = self._get_index(idx)
        self.component.removeByIndex(index, count)

    # endregion TableRowsPartial Overrides

    def append_rows(self, count=1) -> None:
        """
        Appends rows to the table.

        Args:
            count (int, optional): Number of rows to append. Defaults to ``1``.
        """
        self.insert_by_index(-1, count)

    def append_row(self) -> WriteTableRow:
        """
        Appends a row to the table.
        """
        self.append_rows(1)
        return self[-1]
