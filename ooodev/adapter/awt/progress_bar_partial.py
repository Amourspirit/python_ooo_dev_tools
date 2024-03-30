from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XProgressBar

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.color import Color
    from ooodev.utils.type_var import UnoInterface


class ProgressBarPartial:
    """
    Partial class for XProgressBar.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XProgressBar, interface: UnoInterface | None = XProgressBar) -> None:
        """
        Constructor

        Args:
            component (XProgressBar): UNO Component that implements ``com.sun.star.awt.XProgressBar`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XProgressBar``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XProgressBar
    def get_value(self) -> int:
        """
        Gets the current progress value of the progress bar.
        """
        return self.__component.getValue()

    def set_background_color(self, color: Color) -> None:
        """
        Sets the background color of the control.

        Args:
            ~ooodev.utils.color.Color: Color
        """
        self.__component.setBackgroundColor(color)  # type: ignore

    def set_foreground_color(self, color: Color) -> None:
        """
        Sets the foreground color of the control.

        Args:
            ~ooodev.utils.color.Color: Color
        """
        self.__component.setForegroundColor(color)  # type: ignore

    def set_range(self, min_val: int, max_val: int) -> None:
        """
        Sets the minimum and the maximum progress value of the progress bar.

        If the minimum value is greater than the maximum value, the method exchanges the values automatically.
        """
        self.__component.setRange(min_val, max_val)

    def set_value(self, value: int) -> None:
        """
        Sets the progress value of the progress bar.
        """
        self.__component.setValue(value)

    # endregion XProgressBar
