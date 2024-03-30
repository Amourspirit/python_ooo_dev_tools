from __future__ import annotations
from typing import Any, TYPE_CHECKING
import datetime
import uno

from com.sun.star.awt import XDateField

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.date_time_util import DateUtil

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class DateFieldPartial:
    """
    Partial class for XDateField.
    """

    def __init__(self, component: XDateField, interface: UnoInterface | None = XDateField) -> None:
        """
        Constructor

        Args:
            component (XDateField): UNO Component that implements ``com.sun.star.awt.XDateField`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDateField``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XDateField
    def get_date(self) -> datetime.date:
        """
        Gets the date value which is currently displayed in the date field.

        Returns:
            date: The date value.
        """
        return DateUtil.uno_date_to_date(self.__component.getDate())

    def get_first(self) -> datetime.date:
        """
        Gets the currently set first value which is set on POS1 key.

        Returns:
            date: The first value.
        """
        return DateUtil.uno_date_to_date(self.__component.getFirst())

    def get_last(self) -> datetime.date:
        """
        Gets the currently set last value which is set on END key.

        Returns:
            date: The last value.
        """
        return DateUtil.uno_date_to_date(self.__component.getLast())

    def get_max(self) -> datetime.date:
        """
        Gets the currently set maximum date value that can be entered by the user.

        Returns:
            date: The maximum date value.
        """
        return DateUtil.uno_date_to_date(self.__component.getMax())

    def get_min(self) -> datetime.date:
        """
        Gets the currently set minimum date value that can be entered by the user.

        Returns:
            date: The minimum date value.
        """
        return DateUtil.uno_date_to_date(self.__component.getMin())

    def is_empty(self) -> bool:
        """
        Gets whether currently an empty value is set for the date.
        """
        return self.__component.isEmpty()

    def is_long_format(self) -> bool:
        """
        Gets if the long date format is currently used.
        """
        return self.__component.isLongFormat()

    def is_strict_format(self) -> bool:
        """
        Gets whether the format is currently checked during user input.
        """
        return self.__component.isStrictFormat()

    def set_date(self, value: datetime.date) -> None:
        """
        Sets the date value which is displayed in the date field.

        Args:
            value (date): The date value to be set.
        """
        self.__component.setDate(DateUtil.date_to_uno_date(value))

    def set_empty(self) -> None:
        """
        Sets an empty value for the date.
        """
        self.__component.setEmpty()

    def set_first(self, value: datetime.date) -> None:
        """
        Sets the first value to be set on POS1 key.

        Args:
            value (date): The first value to be set.
        """
        self.__component.setFirst(DateUtil.date_to_uno_date(value))

    def set_last(self, value: datetime.date) -> None:
        """
        Sets the last value to be set on END key.

        Args:
            value (date): The last value to be set.
        """
        self.__component.setLast(DateUtil.date_to_uno_date(value))

    def set_long_format(self, long: bool) -> None:
        """
        Determines if the long date format is to be used.
        """
        self.__component.setLongFormat(long)

    def set_max(self, value: datetime.date) -> None:
        """
        Sets the maximum date value that can be entered by the user.

        Args:
            value (date): The maximum date value to be set.
        """
        self.__component.setMax(DateUtil.date_to_uno_date(value))

    def set_min(self, value: datetime.date) -> None:
        """
        Sets the minimum date value that can be entered by the user.

        Args:
            value (date): The minimum date value to be set.
        """
        self.__component.setMin(DateUtil.date_to_uno_date(value))

    def set_strict_format(self, strict: bool) -> None:
        """
        determines if the format is checked during user input.
        """
        self.__component.setStrictFormat(strict)

    # endregion XDateField
