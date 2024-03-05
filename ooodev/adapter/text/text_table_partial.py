from __future__ import annotations
from typing import TYPE_CHECKING, Tuple
import uno

from com.sun.star.text import XTextTable
from ooodev.adapter.text.text_content_partial import TextContentPartial

if TYPE_CHECKING:
    from com.sun.star.text import XTextTableCursor
    from com.sun.star.table import XCell
    from com.sun.star.table import XTableColumns
    from com.sun.star.table import XTableRows
    from ooodev.utils.type_var import UnoInterface


class TextTablePartial(TextContentPartial):
    """
    Partial class for XTextTable.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextTable, interface: UnoInterface | None = XTextTable) -> None:
        """
        Constructor

        Args:
            component (XTextTable): UNO Component that implements ``com.sun.star.text.XTextTable`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextTable``.
        """

        TextContentPartial.__init__(self, component, interface=interface)
        self.__component = component

    # region XTextTable
    def create_cursor_by_cell_name(self, name: str) -> XTextTableCursor:
        """
        Creates a text table cursor and returns the XTextTableCursor interface.

        Initially the cursor is positioned in the cell with the specified name.
        """
        return self.__component.createCursorByCellName(name)

    def get_cell_by_name(self, name: str) -> XCell:
        """
        Example: The cell in the 4th column and third row has the name ``D3``.

        In cells that are split, the naming convention is more complex. In this case the name is a concatenation of the former cell name (i.e. \"D3\") and the number of the new column and row index inside of the original table cell separated by dots.
        This is done recursively.

        Args:
            cell_name (str): The name of the cell.

        Returns:
            XCell: The cell with the specified name.

        Example:
            If the cell ``D3`` is horizontally split, it now contains the cells ``D3.1.1`` and ``D3.1.2``.
        """
        return self.__component.getCellByName(name)

    def get_cell_names(self) -> Tuple[str, ...]:
        """
        Gets the names of all cells in the table.
        """
        return self.__component.getCellNames()

    def get_columns(self) -> XTableColumns:
        """
        Gets the columns of the table.
        """
        return self.__component.getColumns()

    def get_rows(self) -> XTableRows:
        """
        Gets the rows of the table.
        """
        return self.__component.getRows()

    def initialize(self, rows: int, columns: int) -> None:
        """
        Determines the numbers of rows and columns of the text table.

        This method must be called after the object is created and before the object is insert or attached elsewhere.
        """
        self.__component.initialize(rows, columns)

    # endregion XTextTable
