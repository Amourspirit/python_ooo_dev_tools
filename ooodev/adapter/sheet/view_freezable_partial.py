from __future__ import annotations
from typing import Any, TYPE_CHECKING

from com.sun.star.sheet import XViewFreezable

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class ViewFreezablePartial:
    """
    Partial Class for XViewFreezable .

    .. versionadded:: 0.20.0
    """

    def __init__(self, component: XViewFreezable, interface: UnoInterface | None = XViewFreezable) -> None:
        """
        Constructor

        Args:
            component (XViewFreezable): UNO Component that implements ``com.sun.star.sheet.XViewFreezable``.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XViewFreezable``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XViewFreezable
    def freeze_at_position(self, col: int, row: int) -> None:
        """
        Freezes panes with the specified number of columns and rows.

        To freeze only horizontally, specify ``rows`` as 0.
        To freeze only vertically, specify ``col`` as 0.

        Args:
            col (int): The column to freeze.
            row (int): The row to freeze.
        """
        self.__component.freezeAtPosition(col, row)

    def has_frozen_panes(self) -> bool:
        """
        Gets whether the view has frozen panes.
        """
        return self.__component.hasFrozenPanes()

    # endregion XViewFreezable
