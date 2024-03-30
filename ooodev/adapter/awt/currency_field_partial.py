from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XCurrencyField

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class CurrencyFieldPartial:
    """
    Partial class for XCurrencyField.
    """

    def __init__(self, component: XCurrencyField, interface: UnoInterface | None = XCurrencyField) -> None:
        """
        Constructor

        Args:
            component (XCurrencyField): UNO Component that implements ``com.sun.star.awt.XCurrencyField`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XCurrencyField``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XCurrencyField
    def get_decimal_digits(self) -> int:
        """
        Gets the currently set number of decimals.
        """
        return self.__component.getDecimalDigits()

    def get_first(self) -> float:
        """
        Gets the currently set first value which is set on POS1 key.
        """
        return self.__component.getFirst()

    def get_last(self) -> float:
        """
        Gets the currently set last value which is set on END key.
        """
        return self.__component.getLast()

    def get_max(self) -> float:
        """
        Gets the currently set maximum value that can be entered by the user.
        """
        return self.__component.getMax()

    def get_min(self) -> float:
        """
        Gets the currently set minimum value that can be entered by the user.
        """
        return self.__component.getMin()

    def get_spin_size(self) -> float:
        """
        Gets the currently set increment value for the spin button.
        """
        return self.__component.getSpinSize()

    def get_value(self) -> float:
        """
        Gets the value which is currently displayed in the currency field.
        """
        return self.__component.getValue()

    def is_strict_format(self) -> bool:
        """
        Gets whether the format is currently checked during user input.
        """
        return self.__component.isStrictFormat()

    def set_decimal_digits(self, digits: int) -> None:
        """
        Sets the number of decimals.
        """
        self.__component.setDecimalDigits(digits)

    def set_first(self, value: float) -> None:
        """
        sets the first value to be set on POS1 key.
        """
        self.__component.setFirst(value)

    def set_last(self, value: float) -> None:
        """
        Sets the last value to be set on END key.
        """
        self.__component.setLast(value)

    def set_max(self, value: float) -> None:
        """
        Sets the maximum value that can be entered by the user.
        """
        self.__component.setMax(value)

    def set_min(self, value: float) -> None:
        """
        Sets the minimum value that can be entered by the user.
        """
        self.__component.setMin(value)

    def set_spin_size(self, value: float) -> None:
        """
        sets the increment value for the spin button.
        """
        self.__component.setSpinSize(value)

    def set_strict_format(self, strict: bool) -> None:
        """
        Determines if the format is checked during user input.
        """
        self.__component.setStrictFormat(strict)

    def set_value(self, value: float) -> None:
        """
        Sets the value which is displayed in the currency field.
        """
        self.__component.setValue(value)

    # endregion XCurrencyField
