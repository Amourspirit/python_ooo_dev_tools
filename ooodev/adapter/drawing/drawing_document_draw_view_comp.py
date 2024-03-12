from __future__ import annotations
from typing import cast, TYPE_CHECKING
import contextlib
import uno
from ooo.dyn.awt.point import Point

from ooodev.adapter.frame.controller_comp import ControllerComp
from ooodev.adapter.drawing.draw_view_partial import DrawViewPartial
from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.utils.kind.zoom_kind import ZoomKind
from ooodev.utils.data_type.generic_unit_point import GenericUnitPoint
from ooodev.utils.data_type.generic_unit_size_pos import GenericUnitSizePos
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.utils import info as mInfo

if TYPE_CHECKING:
    from com.sun.star.drawing import DrawingDocumentDrawView  # service
    from com.sun.star.lang import XComponent
    from com.sun.star.drawing import XDrawPage


class DrawingDocumentDrawViewComp(
    ControllerComp, DrawViewPartial, PropertySetPartial, PropertyChangeImplement, VetoableChangeImplement
):
    """
    Class for managing DrawingDocumentDrawView Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XComponent) -> None:
        """
        Constructor

        Args:
            component (XComponent): UNO Component that supports ``com.sun.star.drawing.DrawingDocumentDrawView`` service.
        """

        ControllerComp.__init__(self, component)  # type: ignore
        DrawViewPartial.__init__(self, component=component, interface=None)  # type: ignore
        PropertySetPartial.__init__(self, component=component, interface=None)  # type: ignore
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.DrawingDocumentDrawView",)

    # endregion Overrides
    # region Properties

    @property
    def component(self) -> DrawingDocumentDrawView:
        """DrawingDocumentDrawView Component"""
        # override to satisfy documentation and type
        # return cast("DrawingDocumentDrawView", self._ComponentBase__get_component())  # type: ignore
        return cast("DrawingDocumentDrawView", super().component)  # type: ignore

    @property
    def current_page(self) -> XDrawPage:
        """
        This is the drawing page that is currently visible.
        """
        return self.component.CurrentPage

    @current_page.setter
    def current_page(self, value: XDrawPage) -> None:
        self.component.CurrentPage = value

    @property
    def is_layer_mode(self) -> bool:
        """
        If the view is in layer mode, the user can modify the layer of the model of this view in the user interface.
        """
        return self.component.IsLayerMode

    @is_layer_mode.setter
    def is_layer_mode(self, value: bool) -> None:
        self.component.IsLayerMode = value

    @property
    def is_master_page_mode(self) -> bool:
        """
        If the view is in master page mode, the view shows the master pages of this model.
        """
        return self.component.IsMasterPageMode

    @is_master_page_mode.setter
    def is_master_page_mode(self, value: bool) -> None:
        self.component.IsMasterPageMode = value

    @property
    def view_offset(self) -> GenericUnitPoint[UnitMM100, int] | None:
        """
        Gets/Sets the offset from the top left position of the displayed page to the top left position of the view area.

        When setting value can be a ``Point`` or a ``GenericUnitPoint``.

        **optional**

        Returns:
            GenericUnitPoint[UnitMM100, int] | None: The offset from the top left position of the displayed page to the top left position of the view area.

        Hint
            - ``Point`` can be imported from ``ooo.dyn.awt.point``
        """
        with contextlib.suppress(AttributeError):
            p = self.component.ViewOffset
            return GenericUnitPoint(UnitMM100(p.X), UnitMM100(p.Y))
        return None

    @view_offset.setter
    def view_offset(self, value: Point | GenericUnitPoint[UnitMM100, int]) -> None:
        with contextlib.suppress(AttributeError):
            if mInfo.Info.is_instance(value, GenericUnitPoint):
                p = value.get_point()
                self.component.ViewOffset = Point(p.x, p.y)
            else:
                self.component.ViewOffset = value  # type: ignore

    @property
    def visible_area(self) -> GenericUnitSizePos[UnitMM100, int]:
        """
        Gets the area that is currently visible.
        """
        rect = self.component.VisibleArea
        return GenericUnitSizePos(UnitMM100(rect.X), UnitMM100(rect.Y), UnitMM100(rect.Width), UnitMM100(rect.Height))

    @property
    def zoom_type(self) -> ZoomKind:
        """
        Gets/Sets the zoom type for the document.

        Returns:
            ZoomKind: The zoom type.

        Note:
            After setting to value ``ZoomKind.BY_VALUE``, the ``zoom_value`` property should be set to the desired value.

        Hint:
            - ``ZoomKind`` can be imported from ``ooodev.utils.kind.zoom_kind``
        """
        return ZoomKind(self.component.ZoomType)

    @zoom_type.setter
    def zoom_type(self, value: int | ZoomKind) -> None:
        zoom = ZoomKind(int(value))
        if zoom > ZoomKind.PAGE_WIDTH_EXACT:
            self.component.ZoomType = int(ZoomKind.BY_VALUE)
            if zoom == ZoomKind.ZOOM_50_PERCENT:
                self.component.ZoomValue = 50
            elif zoom == ZoomKind.ZOOM_75_PERCENT:
                self.component.ZoomValue = 75
            elif zoom == ZoomKind.ZOOM_150_PERCENT:
                self.component.ZoomValue = 150
            elif zoom == ZoomKind.ZOOM_200_PERCENT:
                self.component.ZoomValue = 200
            else:
                self.component.ZoomValue = 100
        else:
            self.component.ZoomType = zoom.value

    @property
    def zoom_value(self) -> int:
        """
        Gets/Sets the zoom value to use.

        When this value is set ``zoom_type`` is automatically set to ``ZoomKind.BY_VALUE``.
        """
        return self.component.ZoomValue

    @zoom_value.setter
    def zoom_value(self, value: int) -> None:
        by_val = int(ZoomKind.BY_VALUE)
        if self.component.ZoomType != by_val:
            self.component.ZoomType = by_val
        self.component.ZoomValue = value

    # endregion Properties
