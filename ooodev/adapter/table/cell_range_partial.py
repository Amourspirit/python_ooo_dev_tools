from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.table import XCellRange

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.table import XCell
    from ooodev.utils.type_var import UnoInterface


class CellRangePartial:
    """
    Partial Class for XCellRange.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XCellRange, interface: UnoInterface | None = XCellRange) -> None:
        """
        Constructor

        Args:
            component (XCellRange): UNO Component that implements ``com.sun.star.table.XCellRange`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCellRange``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XCellRange
    def get_cell_by_position(self, column: int, row: int) -> XCell:
        """
        Returns a single cell within the range.

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        return self.__component.getCellByPosition(column, row)

    def get_cell_range_by_name(self, rng: str) -> XCellRange:
        """
        Returns a sub-range of cells within the range.

        The sub-range is specified by its name. The format of the range name is dependent of the context of the table.
        In spreadsheets valid names may be ``A1:C5`` or ``$B$2`` or even defined names for cell ranges such as ``MySpecialCell``.
        """
        return self.__component.getCellRangeByName(rng)

    def get_cell_range_by_position(self, left: int, top: int, right: int, bottom: int) -> XCellRange:
        """
        Returns a sub-range of cells within the range.

        Raises:
            com.sun.star.lang.IndexOutOfBoundsException: ``IndexOutOfBoundsException``
        """
        return self.__component.getCellRangeByPosition(left, top, right, bottom)

    # endregion XCellRange
