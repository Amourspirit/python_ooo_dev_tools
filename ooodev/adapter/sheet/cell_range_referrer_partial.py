from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sheet import XCellRangeReferrer

from ooodev.exceptions import ex as mEx
from ooodev.utils import lo as mLo

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
        self.__interface = interface
        self.__validate(component)
        self.__component = component

    def __validate(self, component: Any) -> None:
        """
        Validates the component.

        Args:
            component (Any): The component to be validated.
        """
        if self.__interface is None:
            return
        if not mLo.Lo.is_uno_interfaces(component, self.__interface):
            raise mEx.MissingInterfaceError(self.__interface)

    # region XCellRangeReferrer
    def get_referred_cells(self) -> XCellRange:
        """
        Gets the cell range that is referred to.
        """
        return self.__component.getReferredCells()

    # endregion XCellRangeReferrer
