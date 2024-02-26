from __future__ import annotations
from typing import Any, TYPE_CHECKING, Generic
import uno
from com.sun.star.drawing import XShape

from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.drawing.shape_comp import ShapeComp
from ooodev.adapter.drawing.shape_partial_props import ShapePartialProps
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils import gen_util as gUtil
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.draw.shapes.const import KNOWN_SHAPES
from ooodev.draw.partial.draw_shape_partial import DrawShapePartial
from ooodev.draw.shapes.partial.styled_shape_partial import StyledShapePartial
from ooodev.draw.shapes.shape_base import ShapeBase
from ooodev.draw.shapes.shape_base import _T


if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst


class DrawShape(
    ShapeBase,
    Generic[_T],
    ShapeComp,
    ShapePartialProps,
    PropertyChangeImplement,
    VetoableChangeImplement,
    DrawShapePartial,
    QiPartial,
    PropPartial,
    StylePartial,
    StyledShapePartial,
):
    def __init__(self, owner: _T, component: XShape, lo_inst: LoInst | None = None) -> None:
        ShapeBase.__init__(self, owner=owner, component=component, lo_inst=lo_inst)
        ShapePartialProps.__init__(self, component=component)  # type: ignore
        # QiPartial needs to be before ShapeComp in this case.
        QiPartial.__init__(self, component=component, lo_inst=self.get_lo_inst())
        ShapeComp.__init__(self, component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        DrawShapePartial.__init__(self, component=component, lo_inst=self.get_lo_inst())
        PropPartial.__init__(self, component=component, lo_inst=self.get_lo_inst())
        StylePartial.__init__(self, component=component)
        StyledShapePartial.__init__(self, component=component, lo_inst=self.get_lo_inst())

    # region Overrides
    def _ComponentBase__get_is_supported(self, component: Any) -> bool:
        """
        Gets whether the component supports a service.

        This class also is used for getting shapes in Write Document.
        So, it is possible that the component is not a shape service such as a ``com.sun.star.text.TextFrame``.
        For this reason will verify ``DrawShape`` with ``XShape`` interface.

        Args:
            component (component): UNO Object

        Returns:
            bool: True if the component is XShape the service, otherwise False.
        """
        if component is None:
            return False
        shape = self.qi(XShape)
        return shape is not None

    def get_shape_type(self) -> str:
        """Returns the shape type of ``general``."""
        return "general"

    def _generate_shape_name(self) -> str:
        return f"DrawShape_{gUtil.Util.generate_random_string(10)}"

    def is_know_shape(self) -> bool:
        """
        Returns True if the shape is known.

        Returns:
            bool: ``True`` if the shape is known; Otherwise, ``False``.

        See Also:
            :py:meth:`~.DrawShape.get_known_shape`
        """
        return self.component.getShapeType() in KNOWN_SHAPES

    def get_known_shape(self) -> ShapeBase[_T]:
        """
        The ``DrawShape`` class is a general class for all shapes.
        This means it may not have all the properties of a specific shape.

        This method returns the known shape if the shape is known;
        Otherwise, it returns itself.

        Returns the known shape.

        See Also:
            :py:meth:`~.DrawShape.is_know_shape`
        """
        # avoid circular import
        from ooodev.draw.shapes.partial.shape_factory_partial import ShapeFactoryPartial

        if not self.is_know_shape():
            return self
        factory = ShapeFactoryPartial(self.get_owner(), lo_inst=self.get_lo_inst())

        return factory.shape_factory(self.component)

    # endregion Overrides
