from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from ooo.dyn.table.cell_content_type import CellContentType

from ooodev.mock import mock_g
from ooodev.adapter.text.cell_properties_comp import CellPropertiesComp
from ooodev.adapter.table.cell_partial import CellPartial
from ooodev.adapter.text.text_partial import TextPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.data_type.cell_obj import CellObj

if TYPE_CHECKING:
    from com.sun.star.text import XTextRange
    from com.sun.star.table import XCell
    from ooodev.proto.component_proto import ComponentT
    from ooodev.write.table.write_cell_text_cursor import WriteCellTextCursor
    from ooodev.write.table.write_table_cell_range import WriteTableCellRange


class WriteTableCell(
    WriteDocPropPartial,
    WriteTablePropPartial,
    CellPropertiesComp,
    CellPartial,
    TextPartial,
    LoInstPropsPartial,
    PropPartial,
    StylePartial,
    QiPartial,
):
    """Represents writer table rows."""

    def __init__(self, owner: ComponentT, component: XCell, cell_obj: CellObj) -> None:
        """
        Constructor

        Args:
            component (XCell): UNO object that supports ``com.sun.star.table.XCell`` interface.
        """
        if not isinstance(owner, WriteTablePropPartial):
            raise ValueError("owner must be a WriteTablePropPartial instance.")
        WriteTablePropPartial.__init__(self, obj=owner.write_table)
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        LoInstPropsPartial.__init__(self, lo_inst=self.write_doc.lo_inst)
        CellPropertiesComp.__init__(self, component=component)  # type: ignore
        CellPartial.__init__(self, component=component, interface=None)  # type: ignore
        TextPartial.__init__(self, component=component, interface=None)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        self._owner = owner
        self._cell_obj = cell_obj

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}(cell_name={self.cell_name})"

    def _set_value(self, value: Any) -> None:
        """Set the cell value."""
        if isinstance(value, (float, int)):
            self.set_value(float(value))
        else:
            self.component.String = str(value)  # type: ignore

    def _get_value(self) -> Any:
        """Get the cell value."""
        t = cast(CellContentType, self.get_type())
        if t == CellContentType.EMPTY:
            return None
        if t == CellContentType.VALUE:
            try:
                return float(self.get_value())
            except ValueError:
                return 0.0
        if t == CellContentType.TEXT:
            return self.component.getString()  # type: ignore
        if t == CellContentType.FORMULA:
            return self.get_formula()

    # region get Other Cell

    def _get_cell_other(self, cell_obj: CellObj) -> WriteTableCell:
        """Get the cell object. From Table or Range."""
        # avoiding unnecessary imports for circular imports
        if type(self.owner).__name__ == "WriteTableCellRange":
            src = cast("WriteTableCellRange", self.owner)
        else:
            src = self.write_table
        return src.get_cell(cell_obj)

    def get_cell_left(self) -> WriteTableCell:
        """
        Get the cell to the left.

        Raises:
            IndexError: If cell above is out of range
        """
        return self._get_cell_other(self.cell_obj.left)

    def get_cell_right(self) -> WriteTableCell:
        """
        Get the cell to the right.

        Raises:
            IndexError: If cell above is out of range
        """
        return self._get_cell_other(self.cell_obj.right)

    def get_cell_up(self) -> WriteTableCell:
        """
        Get the cell above.

        Raises:
            IndexError: If cell above is out of range
        """
        return self._get_cell_other(self.cell_obj.up)

    def get_cell_down(self) -> WriteTableCell:
        """
        Get the cell below.

        Raises:
            IndexError: If cell above is out of range
        """
        return self._get_cell_other(self.cell_obj.down)

    # endregion get Other Cell
    # region SimpleTextPartial Overrides

    def create_text_cursor(self) -> WriteCellTextCursor:
        """
        Creates a text cursor to travel in the given range context.

        Cursor can be used to insert text, paragraphs, hyperlinks, and other text content.

        Returns:
            WriteCellTextCursor: Text cursor
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.write.table.write_cell_text_cursor import WriteCellTextCursor

        cursor = self.component.createTextCursor()  # type: ignore
        return WriteCellTextCursor(owner=self, cursor=cursor, lo_inst=self.lo_inst)

    def create_text_cursor_by_range(self, text_position: XTextRange) -> WriteCellTextCursor:
        """
        The initial position is set to ``text_position``.

        Cursor can be used to insert text, paragraphs, hyperlinks, and other text content.

        Args:
            text_position (XTextRange): The initial position of the new text cursor.

        Returns:
            WriteCellTextCursor: The new text cursor.
        """
        # pylint: disable=import-outside-toplevel
        from ooodev.write.table.write_cell_text_cursor import WriteCellTextCursor

        cursor = self.component.createTextCursorByRange(text_position)  # type: ignore
        return WriteCellTextCursor(owner=self, cursor=cursor, lo_inst=self.lo_inst)

    # endregion SimpleTextPartial Overrides

    @property
    def cell_obj(self) -> CellObj:
        """
        Get the cell object.

        Note:
            The ``CellObj`` returned from this property is a sub-range of the parent range or Table.
            This means the ``CellObj`` contains relative values to the parent range or table.
            For this reason the cell name and index **do not** match up with the parent range.
        """
        # if self._cell_obj is None:
        #     self._cell_obj = self.write_table.range_converter.get_cell_obj_from_str(self.cell_name)
        return self._cell_obj

    @property
    def value(self) -> Any:
        """Get the cell value."""
        return self._get_value()

    @value.setter
    def value(self, value: Any) -> None:
        """Set the cell value."""
        self._set_value(value)

    @property
    def owner(self) -> ComponentT:
        """Owner of this component."""
        return self._owner


if mock_g.FULL_IMPORT:
    from ooodev.write.table.write_cell_text_cursor import WriteCellTextCursor
