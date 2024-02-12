from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.loader import lo as mLo

from ooodev.format.inner.partial.position_size.draw.position_partial import PositionPartial
from ooodev.format.inner.partial.position_size.draw.size_partial import SizePartial
from ooodev.format.inner.partial.position_size.draw.protect_partial import ProtectPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from com.sun.star.drawing import XShape


class StyledShapePartial(PositionPartial, ProtectPartial, SizePartial):
    """Partial Class for Shapes that implements Style support."""

    def __init__(self, component: XShape, lo_inst: LoInst | None = None):
        """
        Constructor.

        Args:
            component (XShape): Shape component.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        PositionPartial.__init__(
            self, factory_name=self.__get_position_factory_name(), component=component, lo_inst=lo_inst
        )
        ProtectPartial.__init__(
            self, factory_name=self.__get_protect_factory_name(), component=component, lo_inst=lo_inst
        )
        SizePartial.__init__(self, factory_name=self.__get_size_factory_name(), component=component, lo_inst=lo_inst)

    def __get_position_factory_name(self) -> str:
        return "ooodev.draw.position"

    def __get_protect_factory_name(self) -> str:
        return "ooodev.draw.protect"

    def __get_size_factory_name(self) -> str:
        return "ooodev.draw.size"
