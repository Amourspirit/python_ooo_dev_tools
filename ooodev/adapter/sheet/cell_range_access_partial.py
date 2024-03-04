from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.sheet import XCellRangesAccess

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.table import XCell
    from com.sun.star.table import XCellRange
    from ooodev.utils.type_var import UnoInterface


class CellRangeAccessPartial:
    """
    Partial class for XCellRangesAccess.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XCellRangesAccess, interface: UnoInterface | None = XCellRangesAccess) -> None:
        """
        Constructor

        Args:
            component (XCellRangesAccess): UNO Component that implements ``com.sun.star.container.XCellRangesAccess`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCellRangesAccess``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XCellRangesAccess
    def get_cell_by_position(self, col: int, row: int, idx: int) -> XCell:
        """
        Returns the cell at the specified position.

        Args:
            col (int): The column index of the cell inside the sheet.
            row (int): The row index of the cell inside the sheet.
            idx (int): the sheet index of the sheet inside the document.

        Returns:
            XCell: The single cell within the range.
        """
        return self.__component.getCellByPosition(col, row, idx)

    def get_cell_range_by_position(
        self, start_col: int, start_row: int, end_col: int, end_row: int, idx: int
    ) -> XCellRange:
        """
        Returns the cell range at the specified position.

        Args:
            start_col (int): The column index of the first cell inside the sheet.
            start_row (int): The row index of the first cell inside the sheet.
            end_col (int): The column index of the last cell inside the sheet.
            end_row (int): The row index of the last cell inside the sheet.
            idx (int): The sheet index of the sheet inside the document.

        Returns:
            XCellRange: The cell range.
        """
        return self.__component.getCellRangeByPosition(start_col, start_row, end_col, end_row, idx)

    def get_cell_ranges_by_name(self, name: str) -> tuple[XCellRange, ...]:
        """
        Returns a sub-range of cells within the range.

        The sub-range is specified by its name.
        The format of the range name is dependent of the context of the table.
        In spreadsheets valid names may be ``Sheet1.A1:C5`` or ``$Sheet1.$B$2`` or even defined names for cell ranges such as ``MySpecialCell``.

        Args:
            name (str): The name of the range.

        Returns:
            tuple[XCellRange, ...]: The cell ranges.
        """
        return self.__component.getCellRangesByName(name)

    # endregion XCellRangesAccess
