from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


from ooodev.adapter.drawing.draw_page_comp import DrawPageComp
from ooodev.adapter.drawing.shapes2_partial import Shapes2Partial
from ooodev.adapter.drawing.shapes3_partial import Shapes3Partial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils.partial.qi_partial import QiPartial
from .partial.draw_page_partial import DrawPagePartial

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage

_T = TypeVar("_T", bound="ComponentT")


class GenericDrawPage(
    DrawPagePartial[_T],
    Generic[_T],
    DrawPageComp,
    Shapes2Partial,
    Shapes3Partial,
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
        self.__owner = owner
        DrawPagePartial.__init__(self, owner=self, component=component, lo_inst=self._lo_inst)
        DrawPageComp.__init__(self, component)
        Shapes2Partial.__init__(self, component=component, interface=None)  # type: ignore
        Shapes3Partial.__init__(self, component=component, interface=None)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self._lo_inst)
        StylePartial.__init__(self, component=component)

    def __len__(self) -> int:
        return self.get_count()

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
