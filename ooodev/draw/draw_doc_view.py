from __future__ import annotations
from typing import TYPE_CHECKING
import uno

from ooodev.adapter.drawing.drawing_document_draw_view_comp import DrawingDocumentDrawViewComp
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.draw.partial.draw_doc_prop_partial import DrawDocPropPartial
from ooodev.draw.draw_page import DrawPage

if TYPE_CHECKING:
    from com.sun.star.lang import XComponent
    from ooo.dyn.awt.point import Point
    from ooodev.draw.draw_doc import DrawDoc
    from ooodev.utils.data_type.generic_unit_point import GenericUnitPoint
    from ooodev.units.unit_mm100 import UnitMM100


class DrawDocView(DrawDocPropPartial, LoInstPropsPartial, DrawingDocumentDrawViewComp, QiPartial, ServicePartial):
    """Draw Doc Controller View class. This class is used to manage the view of a Draw document. It is usually accessed via ``DrawDoc.current_controller.``"""

    def __init__(self, owner: DrawDoc, component: XComponent) -> None:
        """
        Constructor

        Args:
            owner (DrawDoc): Draw document.
            component (XComponent): UNO Component that supports ``com.sun.star.drawing.DrawingDocumentDrawView`` service.
        """
        self._owner = owner
        DrawDocPropPartial.__init__(self, obj=owner)
        LoInstPropsPartial.__init__(self, lo_inst=owner.lo_inst)
        DrawingDocumentDrawViewComp.__init__(self, component)
        QiPartial.__init__(self, component=self.component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=self.component, lo_inst=self.lo_inst)

    # region Properties
    # region DrawDocPropPartial Overrides
    @property
    def current_page(self) -> DrawPage:
        """
        This is the drawing page that is currently visible.
        """
        page = self.component.CurrentPage
        return DrawPage(owner=self, component=page, lo_inst=self.lo_inst)

    @current_page.setter
    def current_page(self, value: DrawPage) -> None:
        self.component.CurrentPage = value.component  # type: ignore

    @property
    def view_offset(self) -> GenericUnitPoint[UnitMM100, int]:
        """
        Gets/Sets the offset from the top left position of the displayed page to the top left position of the view area.

        When setting value can be a ``Point`` or a ``GenericUnitPoint``.

        Returns:
            GenericUnitPoint[UnitMM100, int]: The offset from the top left position of the displayed page to the top left position of the view area.

        Hint
            - ``Point`` can be imported from ``ooo.dyn.awt.point``
        """
        return super().view_offset  # type: ignore

    @view_offset.setter
    def view_offset(self, value: Point | GenericUnitPoint[UnitMM100, int]) -> None:
        super().view_offset = value

    # region DrawDocPropPartial Overrides

    # endregion Properties
