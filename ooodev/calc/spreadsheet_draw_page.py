from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


from ooodev.adapter.container.index_access_partial import IndexAccessPartial
from ooodev.adapter.drawing.shapes2_partial import Shapes2Partial
from ooodev.adapter.drawing.shapes3_partial import Shapes3Partial
from ooodev.adapter.sheet.spreadsheet_draw_page_comp import SpreadsheetDrawPageComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.draw.partial.draw_page_partial import DrawPagePartial
from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils.partial.service_partial import ServicePartial
from .calc_forms import CalcForms

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage

_T = TypeVar("_T", bound="ComponentT")


class SpreadsheetDrawPage(
    DrawPagePartial[_T],
    Generic[_T],
    SpreadsheetDrawPageComp,
    IndexAccessPartial,
    Shapes2Partial,
    Shapes3Partial,
    ServicePartial,
    QiPartial,
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
        DrawPagePartial.__init__(self, owner=self, component=component)
        SpreadsheetDrawPageComp.__init__(self, component)
        IndexAccessPartial.__init__(self, component=component, interface=None)  # type: ignore
        Shapes2Partial.__init__(self, component=component, interface=None)  # type: ignore
        Shapes3Partial.__init__(self, component=component, interface=None)  # type: ignore
        ServicePartial.__init__(self, component=component, lo_inst=self._lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self._lo_inst)
        StylePartial.__init__(self, component=component)
        self._forms = None

    def __len__(self) -> int:
        return self.get_count()

    # region Properties
    @property
    def owner(self) -> _T:
        """Component Owner"""
        return self._owner

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
    def forms(self) -> CalcForms:
        """
        Gets the forms of the draw page.
        """
        if self._forms is None:
            self._forms = CalcForms(owner=self, forms=self.component.getForms(), lo_inst=self._lo_inst)  # type: ignore
        return self._forms

    # endregion Properties
