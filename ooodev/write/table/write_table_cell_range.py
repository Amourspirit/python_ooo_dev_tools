from __future__ import annotations
from typing import Any, overload, Generator, TYPE_CHECKING, Tuple, Sequence
import uno
from com.sun.star.lang import IndexOutOfBoundsException

# from ooodev.mock import mock_g
from ooodev.adapter.style.character_properties_partial import CharacterPropertiesPartial
from ooodev.adapter.style.paragraph_properties_partial import ParagraphPropertiesPartial
from ooodev.adapter.table.cell_partial import CellPartial
from ooodev.adapter.text.cell_range_comp import CellRangeComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils import gen_util as mGenUtil
from ooodev.utils.data_type.cell_values import CellValues
from ooodev.utils.data_type.range_obj import RangeObj
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial
from ooodev.write.table.write_table_cell import WriteTableCell

if TYPE_CHECKING:
    from com.sun.star.table import XCellRange
    from com.sun.star.table import CellAddress
    from com.sun.star.table import XCell
    from com.sun.star.table import CellRangeAddress
    from ooodev.utils.data_type.range_values import RangeValues
    from ooodev.proto.component_proto import ComponentT
    from ooodev.utils.data_type.cell_obj import CellObj
    from ooodev.write.table.write_table_row import WriteTableRow


class WriteTableCellRange(
    WriteDocPropPartial,
    WriteTablePropPartial,
    CellRangeComp,
    CellPartial,
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

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}(range={self.range_obj})"

    # region CellRangeDataPartial overrides
    def set_data_array(self, array: Sequence[Sequence[Any]]) -> None:
        """
        Fills the cell range with values from an array.

        The size of the array must be the same as the size of the cell range. Each element of the array must contain a float or a string.

        Warning:
            The size of the array must be the same as the size of the cell range.
            This means when setting table data the table must be the same size as the data.
            When setting a table range the array must be the same size as the range.
        """
        self.component.setDataArray(array)  # type: ignore

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

        return self.get_cell_by_position(col=col_index, row=row_index)

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

    def get_column_range(self, col: int | str, start_row_idx: int = 0) -> WriteTableCellRange:
        """
        Returns a sub-range of cells within the range.

        Args:
            col (int, str): Zero Based column index or a Colum Letter such as ``A``. If integer then can also be a negative value to get from end.
            start_row_idx (int, optional): Start Row Index. Zero Based. Can be negative to get from end. Defaults to ``0``.

        Returns:
            WriteTableCellRange: Range object.
        """
        if isinstance(col, int):
            col_count = self.range_obj.col_count
            index = mGenUtil.Util.get_index(col, col_count)
            col_rng = self.range_obj.get_col(index)
        else:
            col_rng = self.range_obj.get_col(col)

        if start_row_idx > 0:
            if start_row_idx >= col_rng.row_count:
                raise IndexError(f"Index out of range: start_row_idx={start_row_idx}")
            rv = col_rng.get_range_values()
            col_rng = self.write_table.range_converter.rng_from_position(
                col_start=rv.col_start,
                row_start=start_row_idx,
                col_end=rv.col_end,
                row_end=rv.row_end,
            )

        return self.get_cell_range_by_name(str(col_rng))

    def get_row_range(self, row: int) -> WriteTableCellRange:
        """
        Returns a sub-range of cells within the range for a given row.

        Args:
            row (int): Row Index. Zero Based. Can be negative to get from end.

        Returns:
            WriteTableCellRange: Range object.
        """
        row_count = self.range_obj.row_count
        index = mGenUtil.Util.get_index(row, row_count)
        row_rng = self.range_obj.get_row(index)
        return self.get_cell_range_by_name(str(row_rng))

    # endregion Get Row or Column

    # region CellRangePartial Overrides
    def get_cell_by_position(self, col: int, row: int) -> WriteTableCell:
        """
        Returns a single cell within the range.

        Args:
            col (int): Column. Zero Based column index. Can be negative to get from end.
            row (int): Row. Zero Based row index. Can be negative to get from end.

        Raises:
            IndexError: If the index is out of range.

        Returns:
            WriteTableCell: Cell Object.
        """
        # in a sub-range of cells within the range. Cell Names and indexes do not match up.
        # if the origin range is A1:C4 and the sub-range is A2:C2, then the cell at A2 is at column 0, row 0
        # There is no GetCellByName so some conversion is needed.
        try:
            row_count = self.range_obj.row_count
            row_index = mGenUtil.Util.get_index(row, row_count)

            col_count = self.range_obj.col_count
            col_index = mGenUtil.Util.get_index(col, col_count)

            cell_obj = self.write_table.range_converter.get_cell_obj(values=(col_index, row_index))
            return WriteTableCell(
                owner=self,
                component=self.component.getCellByPosition(col, row),
                cell_obj=cell_obj,
            )
        except IndexOutOfBoundsException as e:
            raise IndexError(f"Index out of range. column={col}, row={row}") from e

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
        rv = rng_obj.get_range_values()

        # for some reason the getCellRangeByName always does not work with the range name
        # when I called this method during WriteTableRow.get_row_data() it crashed badly. The bridge was disposed.

        result = WriteTableCellRange(
            owner=self,
            # component=self.component.getCellRangeByName(rng), # can crash the bridge for unknown reasons.
            component=self.component.getCellRangeByPosition(rv.col_start, rv.row_start, rv.col_end, rv.row_end),
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

    def get_row_data(self, idx: int, as_floats: bool = False) -> Tuple[float | str | None, ...]:
        """
        Gets the row data.

        Args:
            idx (int): Index of the row. Zero Based. Can be negative to get from end.
            as_floats (bool, optional): If ``True`` then get all values as floats. If the cell is not a number then it is converted to ``0.0``. Defaults to ``False``.

        Returns:
            Tuple[float | str | None, ...]: Tuple of values.
        """
        range_row = self.get_row_range(idx)
        data = range_row.get_data() if as_floats else range_row.get_data_array()
        return data[0] if len(data) == 1 else data  # type: ignore

    def get_column_data(
        self, idx: int, as_floats: bool = False, start_row_idx: int = 0
    ) -> Tuple[float | str | None, ...]:
        """
        Gets the column data.

        Args:
            idx (int): Index of the column. Zero Based. Can be negative to get from end.
            as_floats (bool, optional): If ``True`` then get all values as floats. If the cell is not a number then it is converted to ``0.0``. Defaults to ``False``.
            start_row_idx (int, optional): Start Row Index. Zero Based. Can be negative to get from end. Defaults to ``0``.

        Returns:
            Tuple[float | str | None, ...]: Tuple of values.
        """
        range_col = self.get_column_range(idx, start_row_idx)

        data_arr = range_col.get_data() if as_floats else range_col.get_data_array()
        data = [row[0] for row in data_arr]
        return tuple(data)

    def get_table_row_index(self, idx: int) -> int:
        """
        Gets the table row index from the range relative index.

        The range index is the index of the row within this range.
        A range can be a subset of the table.
        This method returns the table row index from the range index.

        Args:
            idx (int): Row index within this range. Zero Based. Can be negative to get from end.

        Returns:
            int: Table Row Index. Zero Based.
        """
        row_count = self.range_obj.row_count
        index = mGenUtil.Util.get_index(idx, row_count)
        # get the first cell of the row
        cell = self.get_cell_by_position(0, index)
        cv = self.write_table.range_converter.get_cell_values(cell.cell_name)
        return cv.row

    def get_table_column_index(self, idx: int) -> int:
        """
        Gets the table row index from the range relative index.

        The range index is the index of the row within this range.
        A range can be a subset of the table.
        This method returns the table row index from the range index.

        Args:
            idx (int): Column index within this range. Zero Based. Can be negative to get from end.

        Returns:
            int: Table Row Index. Zero Based.
        """
        col_count = self.range_obj.col_count
        index = mGenUtil.Util.get_index(idx, col_count)
        # get the first cell of the col
        cell = self.get_cell_by_position(index, 0)
        cv = self.write_table.range_converter.get_cell_values(cell.cell_name)
        return cv.col

    def get_table_range_obj(self, col_idx: int, row_idx: int) -> CellObj:
        """
        Gets a cell object that contains column and row where the cell is located in the actual table.

        Args:
            col_idx (int): Column index within this range. Zero Based. Can be negative to get from end.
            row_idx (int): Row index within this range. Zero Based. Can be negative to get from end.

        Returns:
            CellObj: _description_
        """
        col_count = self.range_obj.col_count
        col_index = mGenUtil.Util.get_index(col_idx, col_count)
        row_count = self.range_obj.row_count
        row_index = mGenUtil.Util.get_index(row_idx, row_count)
        cell = self.get_cell_by_position(col_index, row_index)
        return self.write_table.range_converter.get_cell_obj(cell.cell_name)

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
