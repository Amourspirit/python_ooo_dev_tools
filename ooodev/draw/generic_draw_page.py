from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno

from ooodev.adapter.drawing.draw_page_comp import DrawPageComp
from ooodev.adapter.drawing.shapes2_partial import Shapes2Partial
from ooodev.adapter.drawing.shapes3_partial import Shapes3Partial
from ooodev.draw.partial.draw_page_partial import DrawPagePartial
from ooodev.draw.shapes.partial.shape_factory_partial import ShapeFactoryPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils import gen_util as mGenUtil
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage
    from ooodev.draw.shapes.shape_base import ShapeBase
    from ooodev.proto.component_proto import ComponentT

_T = TypeVar("_T", bound="ComponentT")

# ShapeFactoryPartial implements OfficeDocumentPropPartial


class GenericDrawPage(
    DrawPagePartial[_T],
    Generic[_T],
    LoInstPropsPartial,
    ShapeFactoryPartial[_T],
    DrawPageComp,
    Shapes2Partial,
    Shapes3Partial,
    ServicePartial,
    TheDictionaryPartial,
    QiPartial,
    StylePartial,
):
    """
    Represents a draw page.

    Supports index access.

    .. code-block:: python

        shape = doc.slides[0][0] # get a ooodev.draw.shapes.ShapeBase object
    """

    # Draw page does implement XDrawPage, but it show in the API of DrawPage Service.

    def __init__(self, owner: _T, component: XDrawPage, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo

        self.__owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        ShapeFactoryPartial.__init__(self, owner=owner, lo_inst=self.lo_inst)
        DrawPagePartial.__init__(self, owner=self, component=component, lo_inst=self.lo_inst)
        DrawPageComp.__init__(self, component)
        Shapes2Partial.__init__(self, component=component, interface=None)  # type: ignore
        Shapes3Partial.__init__(self, component=component, interface=None)  # type: ignore
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        TheDictionaryPartial.__init__(self)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=component)

    def __len__(self) -> int:
        return self.get_count()

    def __getitem__(self, index: int) -> ShapeBase[_T]:
        idx = mGenUtil.Util.get_index(index, len(self))
        shape = self.component.getByIndex(idx)  # type: ignore
        return self.shape_factory(shape)

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
