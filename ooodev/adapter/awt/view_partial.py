from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from com.sun.star.awt import XView

from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.units.unit_px import UnitPX
from ooodev.units.size_px import SizePX

if TYPE_CHECKING:
    from com.sun.star.awt import XGraphics
    from ooodev.utils.type_var import UnoInterface
    from ooodev.units.unit_obj import UnitT


class ViewPartial:
    """
    Partial class for XView.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XView, interface: UnoInterface | None = XView) -> None:
        """
        Constructor

        Args:
            component (XView): UNO Component that implements ``com.sun.star.awt.XView`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XView``.
        """

        def validate(comp: Any, obj_type: Any) -> None:
            if obj_type is None:
                return
            if not mLo.Lo.is_uno_interfaces(comp, obj_type):
                raise mEx.MissingInterfaceError(obj_type)

        validate(component, interface)
        self.__component = component

    # region XView
    def draw(self, x: int | UnitT, y: int | UnitT) -> None:
        """
        Draws the object at the specified position.

        If the output should be clipped, the caller has to set the clipping region.

        Args:
            x (int | UnitT): X coordinate in device units. If ``int`` then pixel units.
            y (int | UnitT): Y coordinate in device units. If ``int`` then pixel units.
        """
        x_px = UnitPX.from_unit_val(x)
        y_px = UnitPX.from_unit_val(y)
        self.__component.draw(int(x_px), int(y_px))

    def get_graphics(self) -> XGraphics:
        """
        Gets the output device which was set using the method XView.setGraphics().
        """
        return self.__component.getGraphics()

    def get_size(self) -> SizePX:
        """
        Gets the size of the object in device units.

        A device must be set before.

        Returns:
            SizePX: Size of the object in pixel units.
        """
        sz = self.__component.getSize()
        return SizePX.from_unit_val(sz.Width, sz.Height)

    def set_graphics(self, device: XGraphics) -> bool:
        """
        Sets the output device.
        """
        return self.__component.setGraphics(device)

    def set_zoom(self, x: float, y: float) -> None:
        """
        Sets the zoom factor.

        The zoom factor only affects the content of the view, not the size.
        """
        self.__component.setZoom(x, y)

    # endregion XView
