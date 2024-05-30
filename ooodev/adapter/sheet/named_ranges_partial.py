from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.sheet import XNamedRanges
from ooo.dyn.sheet.named_range_flag import NamedRangeFlagEnum
from ooo.dyn.sheet.border import Border

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.adapter.container.name_access_partial import NameAccessPartial


if TYPE_CHECKING:
    from com.sun.star.table import CellAddress
    from com.sun.star.table import CellRangeAddress
    from ooodev.adapter.sheet.named_range_comp import NamedRangeComp
    from ooodev.utils.type_var import UnoInterface


class NamedRangesPartial(NameAccessPartial["NamedRangeComp"]):
    """
    Partial Class for XNamedRanges.
    """

    def __init__(self, component: XNamedRanges, interface: UnoInterface | None = XNamedRanges) -> None:
        """
        Constructor

        Args:
            component (XNamedRanges): UNO Component that implements ``com.sun.star.sheet.XNamedRanges``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XNamedRanges``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        NameAccessPartial.__init__(self, component=component, interface=None)
        self.__component = component

    # region XNamedRanges
    def add_new_by_name(
        self, name: str, content: str, position: CellAddress, type_enum: NamedRangeFlagEnum | int = 0
    ) -> None:
        """
        Adds a new named range to the collection.

        A cell range address is one possible content of a named range.

        Args:
            name (str): The name of the new named range.
            content (str): the formula expression.
            position (CellAddress): Specifies the base address for relative cell references.
            type_enum (int | NamedRangeFlagEnum, optional): A combination of flags that specify the type of a named range. Defaults to ``0``.

        Returns:
            None:

        Note:
            - ``CellAddress`` can be imported from ``com.sun.star.table``.
            - ``NamedRangeFlagEnum`` is a flags enum and can be imported from ``ooo.dyn.sheet.named_range_flag``.
        """
        self.__component.addNewByName(name, content, position, int(type_enum))

    def add_new_from_titles(self, src: CellRangeAddress, border: Border) -> None:
        """
        Creates named cell ranges from titles in a cell range.

        The names for the named ranges are taken from title cells in the top or bottom row,
        or from the cells of the left or right column of the range (depending on the parameter aBorder.
        The named ranges refer to single columns or rows in the inner part of the original range, excluding the labels.

        Example - The source range is ``A1:B3``. The named ranges shall be created using row titles.
        This requires ``Border.TOP`` for the second parameter. The method creates two named ranges.
        The name of the first is equal to the content of cell ``A1`` and contains the range ``$Sheet.$A$2:$A$3`` (excluding the title cell).
        The latter named range is named using cell B1 and contains the cell range address ``$Sheet.$B$2:$B$3``.

        Args:
            src (CellRangeAddress): The cell range containing the titles.
            border (Border): Specifies the border of the cell range that contains the titles.

        Returns:
            None:

        Note:
            - ``CellRangeAddress`` can be imported from ``com.sun.star.table``.
            - ``Border`` is an enum and can be imported from ``ooo.dyn.sheet.border``.
        """
        self.__component.addNewFromTitles(src, border)  # type: ignore

    def output_list(self, out_pos: CellAddress) -> None:
        """
        Writes a list of all named ranges into the document.

        The first column of the list contains the names. The second column contains the contents of the named ranges.

        Returns:
            None:

        Note:
            - ``CellAddress`` can be imported from ``com.sun.star.table``.
        """
        self.__component.outputList(out_pos)

    def remove_by_name(self, name: str) -> None:
        """
        Removes a named range from the collection.
        """
        self.__component.removeByName(name)

    # endregion XNamedRanges
