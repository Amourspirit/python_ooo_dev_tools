from __future__ import annotations
from typing import Any, overload, TYPE_CHECKING, Tuple
import uno
from com.sun.star.lang import IndexOutOfBoundsException

# from ooodev.mock import mock_g
from ooodev.adapter.text.cell_range_comp import CellRangeComp
from ooodev.adapter.style.character_properties_partial import CharacterPropertiesPartial
from ooodev.adapter.style.paragraph_properties_partial import ParagraphPropertiesPartial
from ooodev.adapter.table.cell_partial import CellPartial
from ooodev.adapter.text.text_partial import TextPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.table.write_table_cell import WriteTableCell
from ooodev.utils.data_type.rng.range_converter import RangeConverter
from ooodev.utils.data_type.cell_values import CellValues
from ooodev.utils.data_type.range_obj import RangeObj

if TYPE_CHECKING:
    from com.sun.star.table import XCellRange
    from com.sun.star.table import CellAddress
    from com.sun.star.table import XCell
    from ooodev.proto.component_proto import ComponentT


class WriteTableCellRange(
    WriteDocPropPartial,
    WriteTablePropPartial,
    CellRangeComp,
    CellPartial,
    TextPartial,
    CharacterPropertiesPartial,
    ParagraphPropertiesPartial,
    LoInstPropsPartial,
    PropPartial,
    StylePartial,
    QiPartial,
):
    """Represents writer table rows."""

    def __init__(self, owner: ComponentT, component: XCellRange, range_obj: RangeObj) -> None:
        """
        Constructor

        Args:
            component (TextTableRow): UNO object that supports ``om.sun.star.text.TextTableRow`` service.
        """
        if not isinstance(owner, WriteTablePropPartial):
            raise ValueError("owner must be a WriteTablePropPartial instance.")
        WriteTablePropPartial.__init__(self, obj=owner.write_table)
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        LoInstPropsPartial.__init__(self, lo_inst=self.write_doc.lo_inst)
        CellRangeComp.__init__(self, component=component)  # type: ignore
        CellPartial.__init__(self, component=component, interface=None)  # type: ignore
        TextPartial.__init__(self, component=component, interface=None)  # type: ignore
        CharacterPropertiesPartial.__init__(self, component=component)  # type: ignore
        ParagraphPropertiesPartial.__init__(self, component=component)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)

        self._owner = owner
        self._range_obj = range_obj

    def __getitem__(self, key: Any) -> WriteTableCell:
        """
        Returns the Write Table Cell.

        Args:
            key (Any): Key. can be a Tuple of (col, row) or a string such as "A1" or a CellObj

        See Also:
            - :meth:`get_cell`
        """
        return self.get_cell(key)

    # region CellRangeDataPartial overrides
    def set_data_array(self, array: Tuple[Tuple[Any, ...], ...]) -> None:
        """
        Fills the cell range with values from an array.

        The size of the array must be the same as the size of the cell range. Each element of the array must contain a float or a string.
        """
        self.component.setDataArray(array)

    # endregion CellRangeDataPartial overrides
    # region get_cell()
    @overload
    def get_cell(self, values: Tuple[int, int]) -> WriteTableCell:
        """
        Gets the cell as WriteTableCell from CellValues.

        Args:
            val (Tuple[int, int]): Cell values.
                Tuple of (col, row). Values are Zero Based.

        Returns:
            WriteTableCell: Cell Object.
        """
        ...

    @overload
    def get_cell(self, col: int, row: int) -> WriteTableCell:
        """
        Gets the cell as WriteTableCell from column and row.

        Args:
            col (int): Column. Zero Based column index.
            row (int): Row. Zero Based row index.
            sheet_idx (int, optional): Sheet index that this cell value belongs to. Default is ``-1``.

        Returns:
            WriteTableCell: Cell Object.
        """
        ...

    @overload
    def get_cell(self, addr: CellAddress) -> WriteTableCell:
        """
        Gets the cell as WriteTableCell from a cell address.

        Args:
            addr (CellAddress): Cell Address.

        Returns:
            WriteTableCell: Cell Object.
        """
        ...

    @overload
    def get_cell(self, cell: XCell) -> WriteTableCell:
        """
        Gets the cell as WriteTableCell from a cell.

        Args:
            cell (XCell): Cell.

        Returns:
            WriteTableCell: Cell Object.
        """
        ...

    @overload
    def get_cell(self, val: CellValues) -> WriteTableCell:
        """
        Gets the cell as WriteTableCell from CellValues.

        Args:
            val (CellValues): Cell values.

        Returns:
            WriteTableCell: Cell Object.

        Hint:
            - ``CellValues`` can be imported from ``ooodev.utils.data_type.cell_values``
        """
        ...

    @overload
    def get_cell(self, name: str) -> WriteTableCell:
        """
        Gets the cell as WriteTableCell from a cell name.

        Args:
            name (str): Cell name such as as ``A23`` or ``Sheet1.A23``

        Returns:
            WriteTableCell: Cell Object.
        """
        ...

    def get_cell(self, *args, **kwargs) -> WriteTableCell:
        """Returns a single cell within the range."""
        cell_obj = self.write_table.range_converter.get_cell_obj(*args, **kwargs)
        return self.get_cell_by_position(column=cell_obj.col_obj.index, row=cell_obj.row - 1)

    # endregion get_cell()

    # region CellRangePartial Overrides
    def get_cell_by_position(self, column: int, row: int) -> WriteTableCell:
        """
        Returns a single cell within the range.

        Raises:
            IndexError: If the index is out of range.

        Returns:
            WriteTableCell: Cell Object.
        """
        try:
            return WriteTableCell(owner=self, component=self.component.getCellByPosition(column, row))  # type: ignore
        except IndexOutOfBoundsException as e:
            raise IndexError(f"Index out of range. column={column}, row={row}") from e

    def get_cell_range_by_name(self, rng: str) -> WriteTableCellRange:
        """
        Returns a sub-range of cells within the range.

        The sub-range is specified by its name. The format of the range name is dependent of the context of the table.
        In spreadsheets valid names may be ``A1:C5`` or ``$B$2`` or even defined names for cell ranges such as ``MySpecialCell``.

        Returns:
            WriteTableCell: Cell Object.
        """
        rng_obj = self.write_table.range_converter.rng_from_str(rng)
        return WriteTableCellRange(
            owner=self,
            component=self.component.getCellRangeByName(rng),
            range_obj=rng_obj,
        )

    def get_cell_range_by_position(self, left: int, top: int, right: int, bottom: int) -> WriteTableCellRange:
        """
        Returns a sub-range of cells within the range.

        Raises:
            IndexError: If the index is out of range.

        Returns:
            WriteTableCell: Cell Object.
        """
        try:
            range_obj = self.write_table.range_converter.rng_from_position(
                col_start=left,
                row_start=top,
                col_end=right,
                row_end=bottom,
            )
            return WriteTableCellRange(
                owner=self,
                component=self.component.getCellRangeByPosition(left, top, right, bottom),
                range_obj=range_obj,
            )

        except IndexOutOfBoundsException as e:
            raise IndexError(f"Index out of range: left:{left}, top:{top}, right:{right}, bottom:{bottom}") from e

    # endregion CellRangePartial Overrides

    @property
    def owner(self) -> ComponentT:
        """Owner of this component."""
        return self._owner

    @property
    def range_obj(self) -> RangeObj:
        """Range Object that represents this cell range."""
        return self._range_obj
