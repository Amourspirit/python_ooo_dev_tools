from __future__ import annotations
from typing import TYPE_CHECKING, Generic
import uno

from ooodev.draw.partial.draw_shape_partial import DrawShapePartial
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.drawing.connector_shape_comp import ConnectorShapeComp
from ooodev.adapter.drawing.shape_partial_props import ShapePartialProps
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.draw.shapes.partial.styled_shape_partial import StyledShapePartial
from ooodev.draw.shapes.shape_base import ShapeBase
from ooodev.draw.shapes.shape_base import _T


if TYPE_CHECKING:
    from com.sun.star.drawing import XShape
    from ooodev.loader.inst.lo_inst import LoInst


class ConnectorShape(
    ShapeBase,
    ConnectorShapeComp,
    Generic[_T],
    DrawShapePartial,
    ShapePartialProps,
    QiPartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
    PropPartial,
    StylePartial,
    StyledShapePartial,
):
    def __init__(self, owner: _T, component: XShape, lo_inst: LoInst | None = None) -> None:
        self._owner = owner
        ShapeBase.__init__(self, owner=owner, component=component, lo_inst=lo_inst)
        ConnectorShapeComp.__init__(self, component)
        ShapePartialProps.__init__(self, component=component)  # type: ignore
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        DrawShapePartial.__init__(self, component=component, lo_inst=self.get_lo_inst())
        QiPartial.__init__(self, component=component, lo_inst=self.get_lo_inst())
        PropPartial.__init__(self, component=component, lo_inst=self.get_lo_inst())
        StylePartial.__init__(self, component=component)
        StyledShapePartial.__init__(self, component=component, lo_inst=self.get_lo_inst())

    def get_shape_type(self) -> str:
        """Returns the shape type of ``com.sun.star.drawing.ConnectorShape``."""
        return "com.sun.star.drawing.ConnectorShape"

    def clone(self) -> ConnectorShape[_T]:
        """Clones the shape."""
        shape = self._clone()
        return ConnectorShape[_T](owner=self._owner, component=shape, lo_inst=self.get_lo_inst())
