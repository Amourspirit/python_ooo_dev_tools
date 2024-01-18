from __future__ import annotations
from typing import TYPE_CHECKING, Generic
import uno

from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.drawing.closed_bezier_shape_comp import ClosedBezierShapeComp
from ooodev.adapter.drawing.shape_partial_props import ShapePartialProps
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ..partial.draw_shape_partial import DrawShapePartial
from .shape_base import ShapeBase, _T


if TYPE_CHECKING:
    from com.sun.star.drawing import XShape


class ClosedBezierShape(
    ShapeBase,
    ClosedBezierShapeComp,
    Generic[_T],
    ShapePartialProps,
    DrawShapePartial,
    QiPartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
    PropPartial,
    StylePartial,
):
    def __init__(self, owner: _T, component: XShape) -> None:
        self.__owner = owner
        ShapeBase.__init__(self, owner=owner, component=component)
        ClosedBezierShapeComp.__init__(self, component)
        ShapePartialProps.__init__(self, component=component)  # type: ignore
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        DrawShapePartial.__init__(self, component=component)
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)
        StylePartial.__init__(self, component=component)
