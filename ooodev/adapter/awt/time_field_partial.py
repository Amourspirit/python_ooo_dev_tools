from __future__ import annotations
from typing import Any, TYPE_CHECKING
import datetime
import uno

from com.sun.star.awt import XTimeField

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.utils.date_time_util import DateUtil

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TimeFieldPartial:
    """
    Partial class for XTimeField.
    """

    def __init__(self, component: XTimeField, interface: UnoInterface | None = XTimeField) -> None:
        """
        Constructor

        Args:
            component (XTimeField): UNO Component that implements ``com.sun.star.awt.XTimeField`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTimeField``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTimeField
    def get_first(self) -> datetime.time:
        """
        Gets the currently set first value which is set on POS1 key.

        Returns:
            time: The first value.
        """
        return DateUtil.uno_time_to_time(self.__component.getFirst())

    def get_last(self) -> datetime.time:
        """
        Gets the currently set last value which is set on END key.

        Returns:
            time: The last value.
        """
        return DateUtil.uno_time_to_time(self.__component.getLast())

    def get_max(self) -> datetime.time:
        """
        Gets the currently set maximum time value that can be entered by the user.

        Returns:
            time: The maximum time value.
        """
        return DateUtil.uno_time_to_time(self.__component.getMax())

    def get_min(self) -> datetime.time:
        """
        Gets the currently set minimum time value that can be entered by the user.

        Returns:
            time: The minimum time value.
        """
        return DateUtil.uno_time_to_time(self.__component.getMin())

    def get_time(self) -> datetime.time:
        """
        Gets the time value which is currently displayed in the time field.

        Returns:
            time: The time value.
        """
        return DateUtil.uno_time_to_time(self.__component.getTime())

    def is_empty(self) -> bool:
        """
        Gets whether currently an empty value is set for the time.
        """
        return self.__component.isEmpty()

    def is_strict_format(self) -> bool:
        """
        Gets whether the format is currently checked during user input.
        """
        return self.__component.isStrictFormat()

    def set_empty(self) -> None:
        """
        Sets an empty value for the time.
        """
        self.__component.setEmpty()

    def set_first(self, value: datetime.time) -> None:
        """
        Sets the first value to be set on POS1 key.

        Args:
            value (time): The first value.
        """
        self.__component.setFirst(DateUtil.time_to_uno_time(value))

    def set_last(self, value: datetime.time) -> None:
        """
        Sets the last value to be set on END key.

        Args:
            value (time): The last value.
        """
        self.__component.setLast(DateUtil.time_to_uno_time(value))

    def set_max(self, value: datetime.time) -> None:
        """
        Sets the maximum time value that can be entered by the user.

        Args:
            value (time): The maximum time value.
        """
        self.__component.setMax(DateUtil.time_to_uno_time(value))

    def set_min(self, value: datetime.time) -> None:
        """
        Sets the minimum time value that can be entered by the user.

        Args:
            value (time): The minimum time value.
        """
        self.__component.setMin(DateUtil.time_to_uno_time(value))

    def set_strict_format(self, strict: bool) -> None:
        """
        Determines if the format is checked during user input.
        """
        self.__component.setStrictFormat(strict)

    def set_time(self, value: datetime.time) -> None:
        """
        Sets the time value which is displayed in the time field.

        Args:
            value (time): The time value.
        """
        self.__component.setTime(DateUtil.time_to_uno_time(value))

    # endregion XTimeField
