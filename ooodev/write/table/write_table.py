from __future__ import annotations
from typing import Any, cast, overload, TYPE_CHECKING, Tuple, TypeVar, Generic
import uno
from com.sun.star.lang import IndexOutOfBoundsException

from ooodev.mock import mock_g
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_table_comp import TextTableComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.data_type.range_obj import RangeObj
from ooodev.utils.data_type.rng.range_converter import RangeConverter
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial
from ooodev.write.table.table_column_separators import TableColumnSeparators
from ooodev.write.table.write_table_cell import WriteTableCell
from ooodev.write.table.write_table_cell_range import WriteTableCellRange
from ooodev.write.table.write_text_table_cursor import WriteTextTableCursor
from ooodev.events.partial.events_partial import EventsPartial


if TYPE_CHECKING:
    from com.sun.star.table import CellAddress
    from com.sun.star.table import CellRangeAddress
    from com.sun.star.table import XCell
    from com.sun.star.text import XTextTable
    from ooodev.utils.data_type.cell_obj import CellObj
    from ooodev.proto.component_proto import ComponentT
    from ooodev.utils.data_type.range_values import RangeValues
    from ooodev.utils.data_type.cell_values import CellValues
    from ooodev.write.table.write_table_rows import WriteTableRows
    from ooodev.write.table.write_table_columns import WriteTableColumns
    from ooodev.write.style.direct.table.table_styler import TableStyler

T = TypeVar("T", bound="ComponentT")


class WriteTable(
    Generic[T],
    WriteTablePropPartial,
    EventsPartial,
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
        EventsPartial.__init__(self)
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
        self._style_direct = None

    def __getitem__(self, key: Any) -> WriteTableCell:
        """
        Gets the cell as WriteTableCell.

        Args:
            key (Any): Key. can be a Tuple of (col, row) or a string such as "A1" or a CellObj.

        Returns:
            WriteTableCell: Table Cell Object.

        Example:
            .. code-block:: python

                >>> table = doc.tables[0]
                >>> cell = table["D3"]
                >>> print(cell, cell.value)
                WriteTableCell(cell_name=C4) Sean Connery

        See Also:
            - :meth:`get_cell`
        """
        return self.get_cell(key)

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

    # region get_cell()
    @overload
    def get_cell(self, cell_obj: CellObj) -> WriteTableCell:
        """
        Gets the cell as WriteTableCell.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            WriteTableCell: Cell Object.
        """
        ...

    @overload
    def get_cell(self, values: Tuple[int, int]) -> WriteTableCell:
        """
        Gets the cell as WriteTableCell.

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
        col_index = cell_obj.col_obj.index
        row_index = cell_obj.row - 1

        return self.get_cell_by_position(column=col_index, row=row_index)

    # endregion get_cell()

    # region get Table Cell Range
    @overload
    def get_cell_range(self, cell_obj: CellObj) -> WriteTableCellRange:
        """
        Gets a range Object representing a range.

        Args:
            cell_obj (CellObj): Cell Object.

        Returns:
            WriteTableCellRange: Range object.
        """
        ...

    @overload
    def get_cell_range(self, range_obj: RangeObj) -> WriteTableCellRange:
        """
        Gets a range object. Returns the same object.

        Args:
            range_obj (RangeObj): Range Object

        Returns:
            WriteTableCellRange: Range object.
        """
        ...

    @overload
    def get_cell_range(
        self, col_start: int, row_start: int, col_end: int, row_end: int, sheet_idx: int = ...
    ) -> WriteTableCellRange:
        """
        Gets a range Object representing a range.

        Args:
            col_start (int): Zero-based start column index.
            row_start (int): Zero-based start row index.
            col_end (int): Zero-based end column index.
            row_end (int): Zero-based end row index.
            sheet_idx (int, optional): Zero-based sheet index that this range value belongs to. Default is -1.

        Returns:
            WriteTableCellRange: Range object.
        """
        ...

    @overload
    def get_cell_range(self, addr: CellRangeAddress) -> WriteTableCellRange:
        """
        Gets a range Object representing a range from a cell range address.

        Args:
            addr (CellRangeAddress): Cell Range Address.

        Returns:
            WriteTableCellRange: Range object.
        """
        ...

    @overload
    def get_cell_range(self, rng: RangeValues) -> WriteTableCellRange:
        """
        Gets a range Object representing a range from a cell range address.

        Args:
            rng (RangeValues): Cell Range Values.

        Returns:
            WriteTableCellRange: Range object.

        Hint:
            - ``RangeValues`` can be imported from ``ooodev.utils.data_type.range_values``
        """
        ...

    @overload
    def get_cell_range(self, rng_name: str) -> WriteTableCellRange:
        """
        Gets a range Object representing a range.

        Args:
            rng_name (str): Range as string such as ``Sheet1.A1:C125`` or ``A1:C125``

        Returns:
            WriteTableCellRange: Range object.
        """
        ...

    def get_cell_range(self, *args, **kwargs) -> WriteTableCellRange:
        """
        Gets a range Object representing a range.

        Args:
            rng_name (str): Range as string such as ``Sheet1.A1:C125`` or ``A1:C125``

        Returns:
            WriteTableCellRange: Range object.
        """
        rng = self.range_converter.get_range_obj(*args, **kwargs)

        return self.get_cell_range_by_name(str(rng))

    # endregion get Table Cell Range

    # region TextTablePartial Overrides
    def create_cursor_by_cell_name(self, name: str) -> WriteTextTableCursor:
        """
        Creates a text table cursor and returns the XTextTableCursor interface.

        Initially the cursor is positioned in the cell with the specified name.
        """
        comp = self.component.createCursorByCellName(name)
        return WriteTextTableCursor(owner=self, cursor=comp)

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
        cell_obj = self.range_converter.get_cell_obj_from_str(name)
        return WriteTableCell(
            owner=self,
            component=self.component.getCellByName(name),
            cell_obj=cell_obj,
        )  # type: ignore

    # endregion TextTablePartial Overrides

    # region CellRangePartial Overrides
    def get_cell_by_position(self, column: int, row: int) -> WriteTableCell:
        """
        Returns a single cell within the range.

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        try:
            cell_obj = self.range_converter.get_cell_obj(values=(column, row))
            return WriteTableCell(
                owner=self,
                component=self.component.getCellByPosition(column, row),
                cell_obj=cell_obj,
            )  # type: ignore
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
            range_obj=self.range_converter.get_offset_range_obj(range_obj),
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
                range_obj=self.range_converter.get_offset_range_obj(range_obj),
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
        """
        Get/Sets the name of the table.

        When setting the name, it will be converted to a valid name by replacing spaces with underscores and removing leading and trailing spaces.
        """
        return self.component.getName()

    @name.setter
    def name(self, name: str) -> None:
        s = name.strip().replace(" ", "_")
        self.component.setName(s)

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

    @property
    def style_direct(self) -> TableStyler:
        """
        Direct Cell Styler.

        Returns:
            CellStyler: Character Styler
        """
        if self._style_direct is None:
            # pylint: disable=import-outside-toplevel
            from ooodev.write.style.direct.table.table_styler import TableStyler

            self._style_direct = TableStyler(owner=self, component=self.component)
            self._style_direct.add_event_observers(self.event_observer)
        return self._style_direct

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.write.table.write_table_rows import WriteTableRows
    from ooodev.write.table.write_table_columns import WriteTableColumns
    from ooodev.write.style.direct.table.table_styler import TableStyler
