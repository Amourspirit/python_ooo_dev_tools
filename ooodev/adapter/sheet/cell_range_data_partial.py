from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple, Sequence
import uno
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

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XCellRangeData
    def get_data_array(self) -> Tuple[Tuple[Any, ...], ...]:
        """
        Gets an array from the contents of the cell range.

        Each element of the result contains a float or a string.
        """
        return self.__component.getDataArray()

    def set_data_array(self, array: Sequence[Sequence[Any]]) -> None:
        """
        Fills the cell range with values from an array.

        The size of the array must be the same as the size of the cell range. Each element of the array must contain a float or a string.

        Warning:
            The size of the array must be the same as the size of the cell range.
            This means when setting table data the table must be the same size as the data.
            When setting a table range the array must be the same size as the range.
        """
        self.__component.setDataArray(array)  # type: ignore

    # endregion XCellRangeData
