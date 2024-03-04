from __future__ import annotations
from typing import Any, overload, Generator, TYPE_CHECKING, Tuple
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
from ooodev.utils.data_type.cell_values import CellValues
from ooodev.utils.data_type.range_obj import RangeObj

if TYPE_CHECKING:
    from com.sun.star.table import XCellRange
    from com.sun.star.table import CellAddress
    from com.sun.star.table import XCell
    from com.sun.star.table import CellRangeAddress
    from ooodev.utils.data_type.range_values import RangeValues
    from ooodev.proto.component_proto import ComponentT
    from ooodev.utils.data_type.cell_obj import CellObj


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
        self._parent = None

    def __getitem__(self, key: Any) -> WriteTableCell:
        """
        Returns the Write Table Cell.

        Args:
            key (Any): Key. can be a Tuple of (col, row) or a string such as "A1" or a CellObj.

        Returns:
            WriteTableCell: Table Cell Object.

        Example:
            .. code-block:: python

                >>> table = doc.tables[0]
                >>> rng = table.get_cell_range_by_name("A1:D10")
                >>> cell = rng["D3"]
                >>> print(cell, cell.value)
                WriteTableCell(cell_name=C4) Sean Connery

        See Also:
            - :meth:`get_cell`
        """
        return self.get_cell(key)

    def __iter__(self) -> Generator[WriteTableCell, Any, Any]:
        """
        Iterates through the cells.

        The iteration is done in a column-major order, meaning that the cells are
        iterated over by column, then by row.

        Example:
            .. code-block:: python

                >>> rng = table.get_cell_range_by_name("A1:C3")
                >>> for cell in rng:
                >>> print(cell, cell.value)
                WriteTableCell(cell_name=A1) Title
                WriteTableCell(cell_name=B1) Year
                WriteTableCell(cell_name=C1) Actor
                WriteTableCell(cell_name=A2) Dr. No
                WriteTableCell(cell_name=B2) 1962
                WriteTableCell(cell_name=C2) Sean Connery
                WriteTableCell(cell_name=A3) From Russia with Love
                WriteTableCell(cell_name=B3) 1963
                WriteTableCell(cell_name=C3) Sean Connery
        """
        for cell in self.range_obj:
            yield self.get_cell(cell)

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
        rng = self.write_table.range_converter.get_range_obj(*args, **kwargs)

        return self.get_cell_range_by_name(str(rng))

    # endregion get Table Cell Range

    # region Get Row or Column
    @overload
    def get_column_range(self, col: int) -> WriteTableCellRange:
        """
        Returns a sub-range of cells within the range.

        Args:
            col (int): Column Index. Zero Based.

        Returns:
            WriteTableCellRange: Range object.
        """
        ...

    @overload
    def get_column_range(self, col: str) -> WriteTableCellRange:
        """
        Returns a sub-range of cells within the range.

        Args:
            col (str): Column Letter such as ``A`` or ``B``.

        Returns:
            WriteTableCellRange: Range object.
        """
        ...

    def get_column_range(self, col: int | str) -> WriteTableCellRange:
        """
        Returns a sub-range of cells within the range.

        Args:
            col (int, str): Zero Based column index or a Colum Letter such as ``A``.

        Returns:
            WriteTableCellRange: Range object.
        """
        col_rng = self.range_obj.get_col(col)
        return self.get_cell_range_by_name(str(col_rng))

    def get_row_range(self, row: int) -> WriteTableCellRange:
        """
        Returns a sub-range of cells within the range for a given row.

        Args:
            row (int): Row Index. Zero Based.

        Returns:
            WriteTableCellRange: Range object.
        """
        row_rng = self.range_obj.get_row(row)
        return self.get_cell_range_by_name(str(row_rng))

    # endregion Get Row or Column

    # region CellRangePartial Overrides
    def get_cell_by_position(self, column: int, row: int) -> WriteTableCell:
        """
        Returns a single cell within the range.

        Raises:
            IndexError: If the index is out of range.

        Returns:
            WriteTableCell: Cell Object.
        """
        # in a sub-range of cells within the range. Cell Names and indexes do not match up.
        # if the origin range is A1:C4 and the sub-range is A2:C2, then the cell at A2 is at column 0, row 0
        # There is no GetCellByName so some conversion is needed.
        try:
            cell_obj = self.write_table.range_converter.get_cell_obj(values=(column, row))
            return WriteTableCell(
                owner=self,
                component=self.component.getCellByPosition(column, row),
                cell_obj=cell_obj,
            )
        except IndexOutOfBoundsException as e:
            raise IndexError(f"Index out of range. column={column}, row={row}") from e

    def get_cell_range_by_name(self, rng: str) -> WriteTableCellRange:
        """
        Returns a sub-range of cells within the range.

        The sub-range is specified by its name. The format of the range name is dependent of the context of the table.
        In spreadsheets valid names may be ``A1:C5`` or ``$B$2`` or even defined names for cell ranges such as ``MySpecialCell``.

        Returns:
            WriteTableCell: Cell Object.

        Note:
            This method returns a sub range of cells within the range.
            Sub-ranges cell names and indexes **do not** match up with the parent range.
            If a sub-ranges is ``A3:D6`` then ``sub_rng["A1"]`` is at column 0, row 0 of the sub-range but has a cell name of ``A3``.

        Example:
            .. code-block:: python

                >>> rng = table.get_cell_range_by_name("A1:D10")
                >>> sub1 = rng.get_cell_range_by_name("A3:D6")
                >>> print("Sub1 cell A1")
                >>> print(cell, cell.value)
                Sub1 cell A1
                WriteTableCell(cell_name=A3) From Russia with Love
                >>> sub2 = sub1.get_cell_range_by_name("A4:D5")
                >>> print("Sub1")
                >>> for cell in sub2:
                >>>     print(cell, cell.value)
                Sub2
                WriteTableCell(cell_name=A4) Goldfinger
                WriteTableCell(cell_name=B4) 1964
                WriteTableCell(cell_name=C4) Sean Connery
                WriteTableCell(cell_name=D4) Guy Hamilton
                WriteTableCell(cell_name=A5) Thunderball
                WriteTableCell(cell_name=B5) 1965
                WriteTableCell(cell_name=C5) Sean Connery
                WriteTableCell(cell_name=D5) Terence Young
        """

        # pylint: disable=protected-access

        rng_obj = self.write_table.range_converter.rng_from_str(rng)

        result = WriteTableCellRange(
            owner=self,
            component=self.component.getCellRangeByName(rng),
            range_obj=self.write_table.range_converter.get_offset_range_obj(rng_obj),
        )
        result._parent = self
        return result

    def get_cell_range_by_position(self, left: int, top: int, right: int, bottom: int) -> WriteTableCellRange:
        """
        Returns a sub-range of cells within the range.

        Raises:
            IndexError: If the index is out of range.

        Returns:
            WriteTableCell: Cell Object.

        Note:
            This method returns a sub-range of cells within the range.
            Sub-ranges cell names and indexes **do not** match up with the parent range.
            If a sub-ranges is ``A3:D6`` then ``sub_rng[(0,0)]`` is at column 0, row 0 of the sub-range but has a cell name of ``A3``.
        """
        # pylint: disable=protected-access
        try:
            range_obj = self.write_table.range_converter.rng_from_position(
                col_start=left,
                row_start=top,
                col_end=right,
                row_end=bottom,
            )
            result = WriteTableCellRange(
                owner=self,
                component=self.component.getCellRangeByPosition(left, top, right, bottom),
                range_obj=self.write_table.range_converter.get_offset_range_obj(range_obj),
            )
            result._parent = self
            return result

        except IndexOutOfBoundsException as e:
            raise IndexError(f"Index out of range: left:{left}, top:{top}, right:{right}, bottom:{bottom}") from e

    # endregion CellRangePartial Overrides

    @property
    def owner(self) -> ComponentT:
        """Owner of this component."""
        return self._owner

    @property
    def range_obj(self) -> RangeObj:
        """
        Range Object that represents this cell range.

        Note:
            The ``RangeObj`` returned from this property is a sub-range of the parent range or Table.
            This means the ``RangeObj`` contains relative values to the parent range or table.
            For this reason the cell names and indexes **do not** match up with the parent range.
            If a sub-ranges is ``A3:D6`` then ``sub_rng[(0,0)]`` is at column 0, row 0 of the sub-range but has a cell name of ``A3``.
            The ``RangeObj`` will always start with ``A1`` (column 0 and row 0).
        """
        return self._range_obj

    @property
    def parent(self) -> WriteTableCellRange | None:
        """Parent of this table cell range."""
        return self._parent
