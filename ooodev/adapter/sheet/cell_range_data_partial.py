from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple

from com.sun.star.sheet import XCellRangeData

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class CellRangeDataPartial:
    """
    Partial Class for XCellRangeData.

    .. versionadded:: 0.32.0
    """

    def __init__(self, component: XCellRangeData, interface: UnoInterface | None = XCellRangeData) -> None:
        """
        Constructor

        Args:
            component (XCellRangeData): UNO Component that implements ``com.sun.star.sheet.XCellRangeData``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCellRangeData``.
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

    # region XCellRangeData
    def get_data_array(self) -> Tuple[Tuple[Any, ...], ...]:
        """
        Gets an array from the contents of the cell range.

        Each element of the result contains a float or a string.
        """
        return self.__component.getDataArray()

    def set_data_array(self, array: Tuple[Tuple[Any, ...], ...]) -> None:
        """
        Fills the cell range with values from an array.

        The size of the array must be the same as the size of the cell range. Each element of the array must contain a float or a string.
        """
        self.__component.setDataArray(array)

    # endregion XCellRangeData
