from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooo.dyn.awt.size import Size
from ooo.dyn.awt.point import Point

from ooodev.adapter.drawing.group_shape_comp import GroupShapeComp
from ooodev.adapter.drawing.shape_partial_props import ShapePartialProps
from ooodev.format.inner.style_partial import StylePartial
from ooodev.units.unit_mm import UnitMM
from ooodev.loader import lo as mLo
from ooodev.utils.data_type.generic_unit_point import GenericUnitPoint
from ooodev.utils.data_type.generic_unit_size import GenericUnitSize
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import XShapeGroup
    from ooodev.loader.inst.lo_inst import LoInst


class GroupShape(
    LoInstPropsPartial,
    GroupShapeComp,
    ShapePartialProps,
    QiPartial,
    PropPartial,
    StylePartial,
):
    def __init__(self, component: XShapeGroup, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        GroupShapeComp.__init__(self, component)
        ShapePartialProps.__init__(self, component=component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)

    @property
    def size(self) -> GenericUnitSize[UnitMM, float]:
        """Gets the size of the shape in ``UnitMM`` Values."""
        sz = self.component.getSize()
        return GenericUnitSize(UnitMM.from_mm100(sz.Width), UnitMM.from_mm100(sz.Height))

    @size.setter
    def size(self, value: GenericUnitSize[UnitMM, float]) -> None:
        """Sets the size of the shape in ``UnitMM`` Values."""
        sz = Size(value.width.get_value_mm100(), value.height.get_value_mm100())
        self.set_size(sz)

    @property
    def position(self) -> GenericUnitPoint[UnitMM, float]:
        """Gets the Position of the shape in ``UnitMM`` Values."""
        ps = self.component.getPosition()
        return GenericUnitPoint(UnitMM.from_mm100(ps.X), UnitMM.from_mm100(ps.Y))

    @position.setter
    def position(self, value: GenericUnitPoint[UnitMM, float]) -> None:
        """Sets the Position of the shape in ``UnitMM`` Values."""
        ps = Point(value.x.get_value_mm100(), value.y.get_value_mm100())
        self.set_position(ps)
