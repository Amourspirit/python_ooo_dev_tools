from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.office import draw as mDraw
from ooodev.loader.inst.lo_inst import LoInst

if TYPE_CHECKING:
    from com.sun.star.drawing import XShape


class DrawShapePartial:
    def __init__(self, component: XShape, lo_inst: LoInst | None = None) -> None:
        self.__component = component

    def add_text(self, msg: str, font_size: int = 0, **props) -> None:
        """
        Add text to a shape

        Args:
            msg (str): Text to add
            font_size (int, optional): Font size.
            props (Any, optional): Any extra properties that will be applied to cursor (font) such as ``CharUnderline=1``

        Raises:
            ShapeError: If error occurs.

        Returns:
            None:
        """
        mDraw.Draw.add_text(self.__component, msg, font_size, **props)
