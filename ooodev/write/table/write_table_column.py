from __future__ import annotations
from typing import Any, TYPE_CHECKING, Generator, Tuple

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.utils import gen_util as mGenUtil
from ooodev.utils.data_type.cell_obj import CellObj
from ooodev.utils.data_type.range_obj import RangeObj
from ooodev.utils.data_type.range_values import RangeValues
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial

if TYPE_CHECKING:
    from ooodev.write.table.write_table_cell import WriteTableCell
    from ooodev.write.table.write_table_cell_range import WriteTableCellRange

# Write Table Column Class.
# Write table columns have no components. component.hasElements() is always True;
# However, component.getByIndex(0) always returns None.
# This class has no component but gives access to the cells of the column and the range of the column.


class WriteTableColumn(
    WriteDocPropPartial,
    WriteTablePropPartial,
    EventsPartial,
    LoInstPropsPartial,
):
    """Represents writer table column."""

    def __init__(self, owner: Any, idx: int) -> None:
        """
        Constructor

        Args:
            owner (Any): Owner of this instance.
        """
        if not isinstance(owner, WriteTablePropPartial):
            raise ValueError("owner must be a WriteTablePropPartial instance.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        WriteTablePropPartial.__init__(self, obj=owner.write_table)
        EventsPartial.__init__(self)
        LoInstPropsPartial.__init__(self, lo_inst=owner.write_table.lo_inst)

        self._owner = owner
        self._index = idx
        self._range_obj = None

    def __getitem__(self, key: Any) -> WriteTableCell:
        """
        Returns the Write Table Cell. The cell must exist in the current column.

        Args:
            key (Any): Key. can be a integer such as ``2`` for row index (``-1`` get last cell in col, ``-2`` second last) or a string such as "A1" or a ``CellObj``.

        Raises:
            IndexError: If the key is out of range.

        Returns:
            WriteTableCell: Table Cell Object.

        Note:
            If key is an integer then it is assumed to be a row index.
            If key is a string then it is assumed to be a cell name.

            Cell names and ``CellObj`` are relative to the current column.
            If the current column is the first column of the table then the cell names and ``CellObj`` are the same as the parent table.
            If the column index is 3 then ``col[`A1`]`` is the same as ``table[`C1`]``.

            No mater the column index the first cell of the column is always ``col[0]`` or ``col['A1']``.

        Example:
            .. code-block:: python

                >>> table = doc.tables[0]
                >>> row = table.rows[3]
                >>> cell = row["A1"] # or row[0]
                >>> print(cell, cell.value)
                WriteTableCell(cell_name=A4) Goldfinger
        """
        if isinstance(key, int):
            vals = self._get_range_values()
            index = self._get_index(key)
            cell_obj = CellObj.from_idx(col_idx=vals.col_start, row_idx=index, sheet_idx=vals.sheet_idx)
            # pylint: disable=unsupported-membership-test
            if cell_obj not in self.range_obj:
                raise IndexError(f"Index {key} is out of range.")
            return self.write_table[cell_obj]

        cell_range = self.get_cell_range()
        return cell_range[key]

    def __iter__(self) -> Generator[WriteTableCell, None, None]:
        """Iterates through the cells of the row."""
        # pylint: disable=not-an-iterable
        if self._index < 0:
            raise IndexError("Index is not set.")
        for cell_obj in self.range_obj:
            yield self.write_table[cell_obj]

    def __repr__(self) -> str:
        if self._index < 0:
            return f"WriteTableColumn(index={self.index})"
        return f"WriteTableColumn(index={self.index}, range={self.range_obj})"

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
        count = self.range_obj.col_count
        return mGenUtil.Util.get_index(idx, count, allow_greater)

    def get_cell_range(self) -> WriteTableCellRange:
        """Gets the range of this column."""
        return self.write_table.get_cell_range(self.range_obj)

    def _get_range_values(self) -> RangeValues:
        """Gets the range values of this row."""
        col_start = self.index
        col_end = self.index
        row_start = 0
        row_end = len(self.write_table.rows) - 1

        return RangeValues(col_start=col_start, col_end=col_end, row_start=row_start, row_end=row_end, sheet_idx=-2)

    def get_column_data(self, as_floats: bool = False, start_row_idx: int = 0) -> Tuple[float | str | None, ...]:
        """
        Gets the data of the column.

        Args:
            as_floats (bool, optional): If ``True`` then get all values as floats. If the cell is not a number then it is converted to ``0.0``. Defaults to ``False``.
            start_row_idx (int, optional): Start Row Index. Zero Based. Can be negative to get from end. Defaults to ``0``.

        Returns:
            Tuple[float | str | None, ...]: Column data. If ``as_floats`` is ``True`` then all values are floats.
        """

        cell_range = self.write_table.get_cell_range(self.range_obj)
        return cell_range.get_column_data(idx=0, as_floats=as_floats, start_row_idx=start_row_idx)

    @property
    def owner(self) -> Any:
        """Owner of this instance."""
        return self._owner

    @property
    def index(self) -> int:
        """Index of this column."""
        return self._index

    @property
    def range_obj(self) -> RangeObj:
        """
        Range Object that represents this column cell range.
        """
        if self._range_obj is None:
            if self._index < 0:
                raise IndexError("Index is not set.")
            self._range_obj = RangeObj.from_range(self._get_range_values())
        return self._range_obj
