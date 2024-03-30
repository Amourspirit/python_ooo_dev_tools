from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XLayoutConstrains
from ooo.dyn.awt.size import Size
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.units.size_px import SizePX
from ooodev.units.unit_px import UnitPX

if TYPE_CHECKING:
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.type_var import UnoInterface


class LayoutConstrainsPartial:
    """
    Partial class for XLayoutConstrains.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XLayoutConstrains, interface: UnoInterface | None = XLayoutConstrains) -> None:
        """
        Constructor

        Args:
            component (XLayoutConstrains): UNO Component that implements ``com.sun.star.awt.XLayoutConstrains`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XLayoutConstrains``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XLayoutConstrains
    def calc_adjusted_size(self, width: int | UnitT, height: int | UnitT) -> SizePX:
        """
        Calculates the adjusted size for a given maximum size.

        Args:
            width (int | UnitT): Width. If ``int`` then pixel units.
            height (int | UnitT): Height. If ``int`` then pixel units.

        Returns:
            SizePX: Adjusted size in pixel units.
        """
        width_px = UnitPX.from_unit_val(width)
        height_px = UnitPX.from_unit_val(height)
        sz = Size(int(width_px), int(height_px))

        result = self.__component.calcAdjustedSize(sz)
        return SizePX(UnitPX(result.Width), UnitPX(result.Height))

    def get_minimum_size(self) -> SizePX:
        """
        Gets the minimum size for this component.

        Returns:
            SizePX: Minimum size in pixel units.
        """
        result = self.__component.getMinimumSize()
        return SizePX.from_unit_val(result.Width, result.Height)

    def get_preferred_size(self) -> SizePX:
        """
        Gets the preferred size for this component.

        Returns:
            SizePX: Preferred size in pixel units.
        """
        result = self.__component.getPreferredSize()
        return SizePX.from_unit_val(result.Width, result.Height)

    # endregion XLayoutConstrains
