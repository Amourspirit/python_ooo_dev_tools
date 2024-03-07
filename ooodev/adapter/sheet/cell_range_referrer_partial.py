from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sheet import XCellRangeReferrer

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface
    from com.sun.star.table import XCellRange


class CellRangeReferrerPartial:
    """
    Partial Class for XCellRangeReferrer.

    .. versionadded:: 0.20.0
    """

    def __init__(self, component: XCellRangeReferrer, interface: UnoInterface | None = XCellRangeReferrer) -> None:
        """
        Constructor

        Args:
            component (XCellRangeReferrer): UNO Component that implements ``com.sun.star.sheet.XCellRangeReferrer``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCellRangeReferrer``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XCellRangeReferrer
    def get_referred_cells(self) -> XCellRange:
        """
        Gets the cell range that is referred to.
        """
        return self.__component.getReferredCells()

    # endregion XCellRangeReferrer
