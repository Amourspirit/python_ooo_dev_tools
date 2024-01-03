from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.office import draw as mDraw
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from .generic_draw_page import GenericDrawPage

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage

_T = TypeVar("_T", bound="ComponentT")


class DrawPage(
    GenericDrawPage,
    Generic[_T],
    PropertyChangeImplement,
    VetoableChangeImplement,
    PropPartial,
):
    """Represents a draw page."""

    # Draw page does implement XDrawPage, but it show in the API of DrawPage Service.

    def __init__(self, owner: _T, component: XDrawPage) -> None:
        GenericDrawPage.__init__(self, owner=owner, component=component)
        self.__owner = owner
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)

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

    @property
    def name(self) -> str:
        """
        Gets/Sets the name of the draw page.

        Note:
            Naming for Impress pages seems a little different then Draw pages.
            Attempting to name a Draw page `Slide #` where `#` is a number will fail and Draw will auto name the page.
            It seems that `Slide` followed by a space and a number is reserved for Impress.
        """
        return self.component.Name  # type: ignore

    @name.setter
    def name(self, value: str) -> None:
        self.component.Name = value  # type: ignore

    # endregion Properties
