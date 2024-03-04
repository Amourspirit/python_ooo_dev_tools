from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sheet import XViewSplitable

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.units.unit_obj import UnitT

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ViewSplitablePartial:
    """
    Partial Class for XViewSplitable.

    .. versionadded:: 0.20.0
    """

    def __init__(self, component: XViewSplitable, interface: UnoInterface | None = XViewSplitable) -> None:
        """
        Constructor

        Args:
            component (XViewSplitable): UNO Component that implements ``com.sun.star.sheet.XViewSplitable``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XViewSplitable``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XViewSplitable
    def get_is_window_split(self) -> bool:
        """
        Gets whether the window is split.
        """
        return self.__component.getIsWindowSplit()

    def get_split_column(self) -> int:
        """
        Gets the column where the window is split.
        """
        return self.__component.getSplitColumn()

    def get_split_horizontal(self) -> int:
        """
        Gets the horizontal position in pixels where the view is split.
        """
        return self.__component.getSplitHorizontal()

    def get_split_row(self) -> int:
        """
        Gets the row where the window is split.
        """
        return self.__component.getSplitRow()

    def get_split_vertical(self) -> int:
        """
        Gets the vertical position in pixels where the view is split.
        """
        return self.__component.getSplitVertical()

    def split_at_position(self, x: int | UnitT, y: int | UnitT) -> None:
        """
        Splits the window at the specified position.

        Args:
            x (int, UnitT): The horizontal position in pixels or ``UnitT`` where the view is split.
            y (int, UnitT): The vertical position in pixels or ``UnitT`` where the view is split.
        """
        try:
            x = round(x.get_value_px())  # type: ignore
        except AttributeError:
            x = int(x)  # type: ignore
        try:
            y = round(y.get_value_px())  # type: ignore
        except AttributeError:
            y = int(y)  # type: ignore
        self.__component.splitAtPosition(x, y)

    # endregion XViewSplitable
