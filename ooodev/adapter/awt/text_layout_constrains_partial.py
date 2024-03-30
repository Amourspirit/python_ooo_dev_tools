from __future__ import annotations
from typing import Any, TYPE_CHECKING, Tuple
import uno

from com.sun.star.awt import XTextLayoutConstrains
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.units.size_px import SizePX
from ooodev.units.unit_px import UnitPX

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class TextLayoutConstrainsPartial:
    """
    Partial class for XTextLayoutConstrains.
    """

    # pylint: disable=unused-argument

    def __init__(
        self, component: XTextLayoutConstrains, interface: UnoInterface | None = XTextLayoutConstrains
    ) -> None:
        """
        Constructor

        Args:
            component (XTextLayoutConstrains): UNO Component that implements ``com.sun.star.awt.XTextLayoutConstrains`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XTextLayoutConstrains``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XTextLayoutConstrains
    def get_columns_and_lines(self) -> Tuple[int, int]:
        """
        Returns the ideal number of columns and lines for displaying this text.

        Returns:
            Tuple[int, int]: Number of columns and lines.
        """
        return self.__component.getColumnsAndLines(0, 0)  # type: ignore

    def get_minimum_size(self, cols: int, lines: int) -> SizePX:
        """
        Returns the minimum size for a given number of columns and lines.
        """
        sz = self.__component.getMinimumSize(cols, lines)
        return SizePX(UnitPX(sz.Width), UnitPX(sz.Height))

    # endregion XTextLayoutConstrains
