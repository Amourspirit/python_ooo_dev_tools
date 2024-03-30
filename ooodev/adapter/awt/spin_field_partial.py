from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XSpinField

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from com.sun.star.awt import XSpinListener
    from ooodev.utils.type_var import UnoInterface


class SpinFieldPartial:
    """
    Partial class for XSpinField.
    """

    def __init__(self, component: XSpinField, interface: UnoInterface | None = XSpinField) -> None:
        """
        Constructor

        Args:
            component (XSpinField): UNO Component that implements ``com.sun.star.awt.XSpinField`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XSpinField``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XSpinField
    def add_spin_listener(self, listener: XSpinListener) -> None:
        """
        Registers a listener for spin events.
        """
        self.__component.addSpinListener(listener)

    def down(self) -> None:
        """
        Decreases the value by one step.
        """
        self.__component.down()

    def enable_repeat(self, repeat: bool) -> None:
        """
        Enables/disables automatic repeat mode.
        """
        self.__component.enableRepeat(repeat)

    def first(self) -> None:
        """
        Sets the value to the previously set lower value.
        """
        self.__component.first()

    def last(self) -> None:
        """
        Sets the value to the previously set upper value.
        """
        self.__component.last()

    def remove_spin_listener(self, listener: XSpinListener) -> None:
        """
        Un-registers a listener for spin events.
        """
        self.__component.removeSpinListener(listener)

    def up(self) -> None:
        """
        Increases the value by one step.
        """
        self.__component.up()

    # endregion XSpinField
