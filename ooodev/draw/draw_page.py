from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.office import draw as mDraw
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.units import UnitMM
from .generic_draw_page import GenericDrawPage
from .draw_forms import DrawForms

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage
    from ooodev.units import UnitT

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
        self._forms = None

    def __len__(self) -> int:
        return self.component.getCount()

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

    @property
    def forms(self) -> DrawForms:
        """
        Gets the forms of the draw page.
        """
        if self._forms is None:
            self._forms = DrawForms(owner=self, forms=self.component.getForms())  # type: ignore
        return self._forms

    @property
    def width(self) -> UnitMM:
        """
        Gets the width of the draw page.
        """
        return UnitMM.from_mm100(self.component.Width)

    @width.setter
    def width(self, value: UnitT | float) -> None:
        self.component.Width = mDraw.Draw._get_mm100_obj_from_mm(value, 0).value

    # region    borders
    @property
    def border_left(self) -> UnitMM:
        """
        Gets/Sets the left border of the draw page.
        """
        return UnitMM.from_mm100(self.component.BorderLeft)

    @border_left.setter
    def border_left(self, value: UnitT | float) -> None:
        self.component.BorderLeft = mDraw.Draw._get_mm100_obj_from_mm(value, 0).value

    @property
    def border_right(self) -> UnitMM:
        """
        Gets/Sets the right border of the draw page.
        """
        return UnitMM.from_mm100(self.component.BorderRight)

    @border_right.setter
    def border_right(self, value: UnitT | float) -> None:
        self.component.BorderRight = mDraw.Draw._get_mm100_obj_from_mm(value, 0).value

    @property
    def border_top(self) -> UnitMM:
        """
        Gets/Sets the top border of the draw page.
        """
        return UnitMM.from_mm100(self.component.BorderTop)

    @border_top.setter
    def border_top(self, value: UnitT | float) -> None:
        self.component.BorderTop = mDraw.Draw._get_mm100_obj_from_mm(value, 0).value

    @property
    def border_bottom(self) -> UnitMM:
        """
        Gets/Sets the bottom border of the draw page.
        """
        return UnitMM.from_mm100(self.component.BorderBottom)

    @border_bottom.setter
    def border_bottom(self, value: UnitT | float) -> None:
        self.component.BorderBottom = mDraw.Draw._get_mm100_obj_from_mm(value, 0).value

    # endregion borders
    # endregion Properties
