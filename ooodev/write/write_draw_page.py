from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage

from ooodev.adapter.drawing.generic_draw_page_comp import GenericDrawPageComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.draw.partial.draw_page_partial import DrawPagePartial
from ooodev.adapter.drawing.shapes2_partial import Shapes2Partial
from ooodev.adapter.drawing.shapes3_partial import Shapes3Partial
from .write_forms import WriteForms

_T = TypeVar("_T", bound="ComponentT")


class WriteDrawPage(
    DrawPagePartial[_T],
    Generic[_T],
    GenericDrawPageComp,
    Shapes2Partial,
    Shapes3Partial,
    QiPartial,
    PropPartial,
    StylePartial,
):
    """Represents writer Draw Page."""

    def __init__(self, owner: _T, component: XDrawPage) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XDrawPage): UNO object that supports ``com.sun.star.drawing.GenericDrawPage`` service.
        """
        self.__owner = owner
        DrawPagePartial.__init__(self, owner=self, component=component)
        GenericDrawPageComp.__init__(self, component)  # type: ignore
        Shapes2Partial.__init__(self, component=component, interface=None)  # type: ignore
        Shapes3Partial.__init__(self, component=component, interface=None)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        StylePartial.__init__(self, component=component)
        self._forms = None

    def __len__(self) -> int:
        return self.get_count()

    # region Properties
    @property
    def owner(self) -> _T:
        """Owner of this component."""
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
    def forms(self) -> WriteForms:
        """
        Gets the forms of the draw page.
        """
        if self._forms is None:
            self._forms = WriteForms(owner=self, forms=self.component.getForms())  # type: ignore
        return self._forms

    # endregion Properties
