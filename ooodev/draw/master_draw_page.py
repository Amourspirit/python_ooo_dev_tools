from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.drawing.master_page_comp import MasterPageComp
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.office import draw as mDraw
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils.inst.lo.lo_inst import LoInst
from .partial.draw_page_partial import DrawPagePartial

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage

_T = TypeVar("_T", bound="ComponentT")


class MasterDrawPage(
    DrawPagePartial[_T],
    Generic[_T],
    MasterPageComp,
    PropertyChangeImplement,
    VetoableChangeImplement,
    QiPartial,
    PropPartial,
    StylePartial,
):
    """Represents a draw page."""

    # Draw page does implement XDrawPage, but it show in the API of DrawPage Service.

    def __init__(self, owner: _T, component: XDrawPage, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            self._lo_inst = mLo.Lo.current_lo
        else:
            self._lo_inst = lo_inst
        self._owner = owner
        DrawPagePartial.__init__(self, owner=self, component=component, lo_inst=self._lo_inst)
        MasterPageComp.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        QiPartial.__init__(self, component=component, lo_inst=self._lo_inst)
        PropPartial.__init__(self, component=component, lo_inst=self._lo_inst)
        StylePartial.__init__(self, component=component)

    def get_master_page(self) -> MasterDrawPage[_T]:
        """
        Gets master page

        Raises:
            DrawError: If error occurs.

        Returns:
            MasterDrawPage: Master Page.
        """
        page = mDraw.Draw.get_master_page(self.component)  # type: ignore
        return MasterDrawPage(owner=self._owner, component=page, lo_inst=self._lo_inst)

    def get_notes_page(self) -> MasterDrawPage[_T]:
        """
        Gets the notes page of a slide.

        Each draw page has a notes page.

        Raises:
            DrawPageMissingError: If notes page is ``None``.
            DrawPageError: If any other error occurs.

        Returns:
            MasterDrawPage: Notes Page.

        See Also:
            :py:meth:`~.draw.Draw.get_notes_page_by_index`
        """
        page = mDraw.Draw.get_notes_page(self.component)  # type: ignore
        return MasterDrawPage(owner=self._owner, component=page, lo_inst=self._lo_inst)

    def set_master_footer(self, text: str) -> None:
        """
        Sets master footer text.

        Args:
            text (str): Footer text.

        Raises:
            ShapeMissingError: If unable to find footer shape.
            DrawPageError: If any other error occurs.

        Returns:
            None:
        """
        mDraw.Draw.set_master_footer(self.component, text)  # type: ignore

    # region Properties
    @property
    def owner(self) -> _T:
        """Component Owner"""
        return self._owner

    # endregion Properties
