from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from com.sun.star.drawing import XDrawPage

from ooodev.adapter.drawing.shapes_partial import ShapesPartial

if TYPE_CHECKING:
    from ooodev.utils.type_var import UnoInterface


class DrawPagePartial(ShapesPartial):
    """Partial class for XDrawPage interface."""

    # Does no implement any methods.
    def __init__(self, component: XDrawPage, interface: UnoInterface | None = XDrawPage) -> None:
        """
        Constructor

        Args:
            component (XDrawPage): UNO Component that implements ``com.sun.star.drawing.XDrawPage`` interface.
            interface (UnoInterface, optional): The interface to be validated. Defaults to ``XDrawPage``.
        """
        ShapesPartial.__init__(self, component, interface)
