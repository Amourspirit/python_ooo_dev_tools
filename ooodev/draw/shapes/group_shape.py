from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.adapter.drawing.group_shape_comp import GroupShapeComp
from ooodev.adapter.drawing.shape_partial_props import ShapePartialProps
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import XShapeGroup


class GroupShape(
    GroupShapeComp,
    ShapePartialProps,
    QiPartial,
    PropPartial,
    StylePartial,
):
    def __init__(self, component: XShapeGroup) -> None:
        GroupShapeComp.__init__(self, component)
        ShapePartialProps.__init__(self, component=component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)
        StylePartial.__init__(self, component=component)
