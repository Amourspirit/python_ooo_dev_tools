from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.table import XCell

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooo.dyn.table.cell_content_type import CellContentType
    from ooodev.utils.type_var import UnoInterface


class CellPartial:
    """
    Partial Class for XCell.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XCell, interface: UnoInterface | None = XCell) -> None:
        """
        Constructor

        Args:
            component (XCell): UNO Component that implements ``com.sun.star.table.XCell`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCell``.
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

    # region XCell
    def get_error(self) -> int:
        """
        returns the error value of the cell.

        If the cell does not contain a formula, the error is always zero.
        """
        return self.__component.getError()

    def get_formula(self) -> str:
        """
        returns the formula string of a cell.

        Even if the cell does not contain a formula, an assignment of this attribute's value to another cell's formula attribute would create the same cell content. This is because this attribute contains the original text value of a string cell. The value of a value cell will be formatted using the number format's default format or the formula string, including \"=\", of a formula cell.
        """
        return self.__component.getFormula()

    def get_type(self) -> CellContentType:
        """
        Gets the type of the cell.

        Returns:
            CellContentType: The type of the cell.

        Hint:
            - ``CellContentType`` can be imported from ``ooo.dyn.table.cell_content_type``.
        """
        return self.__component.getType()  # type: ignore

    def get_value(self) -> float:
        """
        Gets the floating point value of the cell.

        For a value cell the value is returned, for a string cell zero is returned and for a formula cell the result value of a formula is returned.
        """
        return self.__component.getValue()

    def set_formula(self, formula: str) -> None:
        """
        Sets a formula into the cell.

        When assigned, the string will be interpreted and a value, text or formula cell is created, depending on the text and the number format.
        """
        self.__component.setFormula(formula)

    def set_value(self, value: float) -> None:
        """
        Sets a floating point value into the cell.

        After a call to this method the type of the cell is C``ellContentType.VALUE``.
        """
        self.__component.setValue(value)

    # endregion XCell
