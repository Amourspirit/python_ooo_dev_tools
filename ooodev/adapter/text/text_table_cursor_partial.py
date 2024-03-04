from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.text import XTextTableCursor

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TextTableCursorPartial:
    """
    Partial class for XTextTableCursor.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTextTableCursor, interface: UnoInterface | None = XTextTableCursor) -> None:
        """
        Constructor

        Args:
            component (XTextTableCursor): UNO Component that implements ``com.sun.star.text.XTextTableCursor`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextTableCursor``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTextTableCursor
    def get_range_name(self) -> str:
        """
        The name is the cell name of the top left table cell of the range concatenated by ``:`` with the table cell name of the bottom left table cell of the cell range.
        If the range consists of one table cell only then the name of that table cell is returned.
        """
        return self.__component.getRangeName()

    def go_down(self, count: int, expand: bool = False) -> bool:
        """
        Moves the cursor to the bottom neighbor cell.

        Args:
            count (int): The number of cells to move.
            expand (bool, optional): Specifies whether the selection is to be expanded. Defaults to ``False``.

        Returns:
            bool: ``True`` if the cursor was moved, ``False`` otherwise.
        """
        return self.__component.goDown(count, expand)

    def go_left(self, count: int, expand: bool = False) -> bool:
        """
        Moves the cursor to the left neighbor.

        Args:
            count (int): The number of cells to move.
            expand (bool, optional): Specifies whether the selection is to be expanded. Defaults to ``False``.

        Returns:
            bool: ``True`` if the cursor was moved, ``False`` otherwise.
        """
        return self.__component.goLeft(count, expand)

    def go_right(self, count: int, expand: bool = False) -> bool:
        """
        Moves the cursor to the right neighbor.

        Args:
            count (int): The number of cells to move.
            expand (bool, optional): Specifies whether the selection is to be expanded. Defaults to ``False``.

        Returns:
            bool: ``True`` if the cursor was moved, ``False`` otherwise.
        """
        return self.__component.goRight(count, expand)

    def go_up(self, count: int, expand: bool = False) -> bool:
        """
        Moves the cursor to the top neighbor.

        Args:
            count (int): The number of cells to move.
            expand (bool, optional): Specifies whether the selection is to be expanded. Defaults to ``False``.

        Returns:
            bool: ``True`` if the cursor was moved, ``False`` otherwise.
        """
        return self.__component.goUp(count, expand)

    def goto_cell_by_name(self, cell_name: str, expand: bool = False) -> bool:
        """
        Moves the cursor to the cell with the specified name.

        Args:
            cell_name (str): The name of the cell.
            expand (bool, optional): Specifies whether the selection is to be expanded. Defaults to ``False``.

        Returns:
            bool: ``True`` if the cursor was moved, ``False`` otherwise.
        """
        return self.__component.gotoCellByName(cell_name, expand)

    def goto_end(self, expand: bool = False) -> None:
        """
        Moves the cursor to the bottom right cell of the table.

        Args:
            expand (bool, optional): Specifies whether the selection is to be expanded. Defaults to ``False``.

        Returns:
            None:
        """
        self.__component.gotoEnd(expand)

    def goto_start(self, expand: bool = False) -> None:
        """
        Moves the cursor to the top left cell of the table.

        Args:
            expand (bool, optional): Specifies whether the selection is to be expanded. Defaults to ``False``.

        Returns:
            None:
        """
        self.__component.gotoStart(expand)

    def merge_range(self) -> bool:
        """
        Merges the selected range of cells.

        Returns:
            bool: ``True`` if the range was merged, ``False`` otherwise.
        """
        return self.__component.mergeRange()

    def split_range(self, count: int, horizontal: bool = True) -> bool:
        """
        Splits the range of cells.

        Args:
            count (int): The number of cells to split.
            horizontal (bool, optional): ``True`` if the range should be split vertically. Otherwise it will be split horizontally. Defaults to ``True``.

        Returns:
            bool: ``True`` if the range was split, ``False`` otherwise.
        """
        return self.__component.splitRange(count, horizontal)

    # endregion XTextTableCursor
