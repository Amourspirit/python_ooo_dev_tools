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
            key (key, int): The index of the cell. When getting by index can be a negative value to get from the end.

        Returns:
            WriteForm: The form with the specified index or name.

        See Also:
            - :py:meth:`~ooodev.write.table.write_table_rows.get_by_index`
        """
        return self.get_by_index(key)

    # region IndexAccessPartial overrides
    def get_by_index(self, idx: int) -> WriteTableRow:
        """
        Gets the row at the specified index.

        Args:
            idx (int): Index of row.

        Returns:
            Any: Row at the specified index.
        """
        idx = self._get_index(idx)
        return WriteTableRow(owner=self, component=self.component.getByIndex(idx))

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
