from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


if TYPE_CHECKING:
    from com.sun.star.drawing import XShapes

from ooodev.adapter.drawing.generic_draw_page_comp import GenericDrawPageComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial

T = TypeVar("T", bound="ComponentT")


class WriteDrawPage(Generic[T], GenericDrawPageComp, QiPartial, PropPartial, StylePartial):
    """Represents writer Draw Page."""

    def __init__(self, owner: T, component: XShapes) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XShapes): UNO object that supports ``com.sun.star.drawing.GenericDrawPage`` service.
        """
        self.__owner = owner
        GenericDrawPageComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        StylePartial.__init__(self, component=component)

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
