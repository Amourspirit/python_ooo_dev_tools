from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, TypeVar, Generic
import uno
from com.sun.star.lang import IndexOutOfBoundsException

from ooodev.mock import mock_g
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_table_comp import TextTableComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write import write_text_portions as mWriteTextPortions
from ooodev.write.table.table_column_separators import TableColumnSeparators
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial
from ooodev.write.table.write_table_cell import WriteTableCell
from ooodev.write.table.write_table_cell_range import WriteTableCellRange
from ooodev.utils.data_type.rng.range_converter import RangeConverter

if TYPE_CHECKING:
    from com.sun.star.text import XTextTable
    from com.sun.star.container import XEnumerationAccess
    from ooodev.proto.component_proto import ComponentT
    from ooodev.write.table.write_table_rows import WriteTableRows
    from ooodev.write.table.write_table_columns import WriteTableColumns

T = TypeVar("T", bound="ComponentT")


class WriteTable(
    Generic[T],
    WriteTablePropPartial,
    LoInstPropsPartial,
    WriteDocPropPartial,
    TextTableComp,
    QiPartial,
    StylePartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
):
    """Represents writer text content."""

    # this class can be used to wrap the table created by
    # ooodev.office.write.Write.add_table() method.

    # https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1TextTable.html

    def __init__(self, owner: T, component: XTextTable, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextTable): UNO object that supports ``com.sun.star.text.TextContent`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        WriteTablePropPartial.__init__(self, obj=self)
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        TextTableComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=component, trigger_args=generic_args)  # type: ignore
        VetoableChangeImplement.__init__(self, component=component, trigger_args=generic_args)  # type: ignore
        self._cols = None
        self._rows = None
        self._range_converter = None

    def __getitem__(self, _val: Any) -> WriteTableCell:
        if isinstance(_val, tuple):
            return self.get_cell_by_position(int(_val[0]), int(_val[1]))
        return self.get_cell_by_name(_val)

    def get_text_portions(self) -> mWriteTextPortions.WriteTextPortions[T]:
        """Returns the text portions of this paragraph."""
        return mWriteTextPortions.WriteTextPortions(
            owner=self.owner, component=cast("XEnumerationAccess", self.component), lo_inst=self.lo_inst
        )

    # region TextTablePartial overrides
    def get_columns(self) -> WriteTableColumns:
        """
        Gets the columns of the table.
        """
        return self.columns

    def get_rows(self) -> WriteTableRows:
        """
        Gets the rows of this table.

        Returns:
            WriteTableRows: Table Rows
        """
        return self.rows

    # endregion TextTablePartial overrides

    # region TextTablePartial Overrides
    def get_cell_by_name(self, name: str) -> WriteTableCell:
        """
        Returns the cell with the specified name.

        The cell in the 4th column and third row has the name ``D3``.

        Args:
            name (str): The name of the cell.

        Returns:
            WriteTableCell: The cell with the specified name.

        Note:
            In cells that are split, the naming convention is more complex.
            In this case the name is a concatenation of the former cell name (i.e. ``D3``) and
            the number of the new column and row index inside of the original table cell separated by dots.
            This is done recursively.

            For example, if the cell ``D3`` is horizontally split, it now contains the cells ``D3.1.1`` and ``D3.1.2``.
        """
        return WriteTableCell(owner=self, component=self.component.getCellByName(name))  # type: ignore

    # endregion TextTablePartial Overrides

    # region CellRangePartial Overrides
    def get_cell_by_position(self, column: int, row: int) -> WriteTableCell:
        """
        Returns a single cell within the range.

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
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
        """
        range_obj = self.range_converter.rng_from_str(rng)
        return WriteTableCellRange(
            owner=self,
            component=self.component.getCellRangeByName(rng),
            range_obj=range_obj,
        )

    def get_cell_range_by_position(self, left: int, top: int, right: int, bottom: int) -> WriteTableCellRange:
        """
        Returns a sub-range of cells within the range.

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        try:
            range_obj = self.range_converter.rng_from_position(
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

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    @property
    def name(self) -> str:
        return self.component.getName()

    @name.setter
    def name(self, name: str) -> None:
        self.component.setName(name)

    @property
    def columns(self) -> WriteTableColumns:
        """Table Rows"""
        if self._cols is None:
            # pylint: disable=import-outside-toplevel
            from ooodev.write.table.write_table_columns import WriteTableColumns

            self._cols = WriteTableColumns(owner=self, component=self.component.getColumns())
        return self._cols

    @property
    def rows(self) -> WriteTableRows:
        """Table Rows"""
        if self._rows is None:
            # pylint: disable=import-outside-toplevel
            from ooodev.write.table.write_table_rows import WriteTableRows

            self._rows = WriteTableRows(owner=self, component=self.component.getRows())
        return self._rows

    @property
    def table_column_separators(self) -> TableColumnSeparators:
        """
        Table Column Separators

        Moving a column separator will change the width of the column to the left and right of the separator.
        There is one less separator than there are columns.

        To get the number of separators, use the length of this property (``len(instance.table_column_separators)``).

        The position of each separator must be greater then the previous separator and less then the next separator.
        The position of the last separator must be less then the value of ``table_column_relative_sum`` property.

        Example:
            .. code-block:: python

                table = doc.tables[0]
                table.table_column_separators[1].position = 5312

                for sep in table.table_column_separators:
                    print(sep)
        """
        return TableColumnSeparators(self.component)

    @property
    def table_column_relative_sum(self) -> int:
        """Gets the sum of the relative widths of all columns."""
        return self.component.TableColumnRelativeSum

    @property
    def range_converter(self) -> RangeConverter:
        """Gets access to a range converter."""
        if self._range_converter is None:
            self._range_converter = RangeConverter(lo_inst=self.lo_inst)
        return self._range_converter

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.write.table.write_table_rows import WriteTableRows
    from ooodev.write.table.write_table_columns import WriteTableColumns
