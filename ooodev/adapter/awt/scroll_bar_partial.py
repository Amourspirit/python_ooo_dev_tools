from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XScrollBar
from ooo.dyn.awt.scroll_bar_orientation import ScrollBarOrientationEnum

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.units.unit_px import UnitPX

if TYPE_CHECKING:
    from com.sun.star.awt import XAdjustmentListener
    from ooodev.utils.type_var import UnoInterface
    from ooodev.units.unit_obj import UnitT


class ScrollBarPartial:
    """
    Partial class for XScrollBar.
    """

    def __init__(self, component: XScrollBar, interface: UnoInterface | None = XScrollBar) -> None:
        """
        Constructor

        Args:
            component (XScrollBar): UNO Component that implements ``com.sun.star.awt.XScrollBar`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XScrollBar``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XScrollBar
    def add_adjustment_listener(self, listener: XAdjustmentListener) -> None:
        """
        Registers an adjustment event listener.
        """
        self.__component.addAdjustmentListener(listener)

    def get_block_increment(self) -> int:
        """
        Gets the currently set increment for a block move.
        """
        return self.__component.getBlockIncrement()

    def get_line_increment(self) -> int:
        """
        Gets the currently set increment for a single line move.
        """
        return self.__component.getLineIncrement()

    def get_maximum(self) -> int:
        """
        Gets the currently set maximum scroll value of the scroll bar.
        """
        return self.__component.getMaximum()

    def get_orientation(self) -> ScrollBarOrientationEnum:
        """
        Gets the currently set ScrollBarOrientation of the scroll bar.
        """
        return ScrollBarOrientationEnum(self.__component.getOrientation())

    def get_value(self) -> int:
        """
        Gets the current scroll value of the scroll bar.
        """
        return self.__component.getValue()

    def get_visible_size(self) -> UnitPX:
        """
        returns the currently visible size of the scroll bar.
        """
        return UnitPX(self.__component.getVisibleSize())

    def remove_adjustment_listener(self, listener: XAdjustmentListener) -> None:
        """
        Un-registers an adjustment event listener.
        """
        self.__component.removeAdjustmentListener(listener)

    def set_block_increment(self, n: int) -> None:
        """
        Sets the increment for a block move.
        """
        self.__component.setBlockIncrement(n)

    def set_line_increment(self, n: int) -> None:
        """
        Sets the increment for a single line move.
        """
        self.__component.setLineIncrement(n)

    def set_maximum(self, n: int) -> None:
        """
        Sets the maximum scroll value of the scroll bar.
        """
        self.__component.setMaximum(n)

    def set_orientation(self, n: int | ScrollBarOrientationEnum) -> None:
        """
        Sets the ScrollBarOrientation of the scroll bar.

        - 0: ``HORIZONTAL``
        - 1: ``VERTICAL``

        Args:
            n (int | ScrollBarOrientationEnum): ScrollBarOrientation value.
        """
        self.__component.setOrientation(int(n))

    def set_value(self, n: int) -> None:
        """
        Sets the scroll value of the scroll bar.
        """
        self.__component.setValue(n)

    def set_values(self, val: int, visible: int, max_scroll: int) -> None:
        """
        Sets the scroll value, visible area and maximum scroll value of the scroll bar.
        """
        self.__component.setValues(val, visible, max_scroll)

    def set_visible_size(self, size: int | UnitT) -> None:
        """
        Sets the visible size of the scroll bar.
        """
        n = UnitPX.from_unit_val(size)
        self.__component.setVisibleSize(int(n))

    # endregion XScrollBar
