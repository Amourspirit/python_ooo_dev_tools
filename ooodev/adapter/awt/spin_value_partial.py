from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XSpinValue

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XAdjustmentListener
    from ooodev.utils.type_var import UnoInterface


class SpinValuePartial:
    """
    Partial class for XSpinValue.
    """

    def __init__(self, component: XSpinValue, interface: UnoInterface | None = XSpinValue) -> None:
        """
        Constructor

        Args:
            component (XSpinValue): UNO Component that implements ``com.sun.star.awt.XSpinValue`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSpinValue``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XSpinValue
    def add_adjustment_listener(self, listener: XAdjustmentListener) -> None:
        """
        Registers an adjustment event listener.
        """
        self.__component.addAdjustmentListener(listener)

    def get_maximum(self) -> int:
        """
        Gets the currently set maximum value of the control
        """
        return self.__component.getMaximum()

    def get_minimum(self) -> int:
        """
        returns the currently set minimum value of the control
        """
        return self.__component.getMinimum()

    def get_orientation(self) -> int:
        """
        returns the current orientation of the control
        """
        return self.__component.getOrientation()

    def get_spin_increment(self) -> int:
        """
        Returns the value by which the current value of the control should be incremented or decremented upon spinning.
        """
        return self.__component.getSpinIncrement()

    def get_value(self) -> int:
        """
        returns the current value of the control.
        """
        return self.__component.getValue()

    def remove_adjustment_listener(self, listener: XAdjustmentListener) -> None:
        """
        Un-registers an adjustment event listener.
        """
        self.__component.removeAdjustmentListener(listener)

    def set_maximum(self, max_value: int) -> None:
        """
        Sets the maximum value which can be set on the control
        """
        self.__component.setMaximum(max_value)

    def set_minimum(self, min_value: int) -> None:
        """
        Sets the minimum value which can be set on the control
        """
        self.__component.setMinimum(min_value)

    def set_orientation(self, orientation: int) -> None:
        """
        controls the orientation of the control

        Raises:
            com.sun.star.lang.NoSupportException: ``NoSupportException``
        """
        self.__component.setOrientation(orientation)

    def set_spin_increment(self, spin_increment: int) -> None:
        """
        sets the value by which the current value of the control should be incremented or decremented upon spinning.
        """
        self.__component.setSpinIncrement(spin_increment)

    def set_value(self, value: int) -> None:
        """
        Sets the current value of the control
        """
        self.__component.setValue(value)

    def set_values(self, min_value: int, max_value: int, current_value: int) -> None:
        """
        Sets the value and value range of the control
        """
        self.__component.setValues(min_value, max_value, current_value)

    # endregion XSpinValue
