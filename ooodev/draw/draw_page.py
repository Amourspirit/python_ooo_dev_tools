from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.container.index_access_partial import IndexAccessPartial
from ooodev.adapter.drawing.shapes2_partial import Shapes2Partial
from ooodev.adapter.drawing.shapes3_partial import Shapes3Partial
from ooodev.adapter.drawing.draw_page_comp import DrawPageComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import draw as mDraw
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from .partial.draw_page_partial import DrawPagePartial

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage

_T = TypeVar("_T", bound="ComponentT")


class DrawPage(
    DrawPagePartial[_T],
    Generic[_T],
    DrawPageComp,
    IndexAccessPartial,
    Shapes2Partial,
    Shapes3Partial,
    PropertyChangeImplement,
    VetoableChangeImplement,
    QiPartial,
    PropPartial,
    StylePartial,
):
    """Represents a draw page."""

    # Draw page does implement XDrawPage, but it show in the API of DrawPage Service.

    def __init__(self, owner: _T, component: XDrawPage) -> None:
        self.__owner = owner
        DrawPagePartial.__init__(self, owner=self, component=component)
        DrawPageComp.__init__(self, component)
        IndexAccessPartial.__init__(self, component=component, interface=None)  # type: ignore
        Shapes2Partial.__init__(self, component=component, interface=None)  # type: ignore
        Shapes3Partial.__init__(self, component=component, interface=None)  # type: ignore
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)
        StylePartial.__init__(self, component=component)

    def get_master_page(self) -> DrawPage[_T]:
        """
        Gets master page

        Raises:
            DrawError: If error occurs.

        Returns:
            DrawPage: Master Page.
        """
        page = mDraw.Draw.get_master_page(self.component)  # type: ignore
        return DrawPage(self.__owner, page)

    def get_notes_page(self) -> DrawPage[_T]:
        """
        Gets the notes page of a slide.

        Each draw page has a notes page.

        Raises:
            DrawPageMissingError: If notes page is ``None``.
            DrawPageError: If any other error occurs.

        Returns:
            DrawPage: Notes Page.

        See Also:
            :py:meth:`~.draw.Draw.get_notes_page_by_index`
        """
        page = mDraw.Draw.get_notes_page(self.component)  # type: ignore
        return DrawPage(self.__owner, page)

    # region Properties
    @property
    def owner(self) -> _T:
        """Component Owner"""
        return self.__owner

    # endregion Properties
