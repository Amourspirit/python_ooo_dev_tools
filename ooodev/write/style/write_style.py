from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno

if TYPE_CHECKING:
    from com.sun.star.style import XStyle


from ooodev.adapter.style.style_comp import StyleComp
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial

T = TypeVar("T", bound="ComponentT")


class WriteStyle(Generic[T], StyleComp, QiPartial, PropPartial):
    """Represents writer Style."""

    def __init__(self, owner: T, component: XStyle) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XStyle): UNO object that supports ``com.sun.star.style.Style`` service.
        """
        self.__owner = owner
        StyleComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    # region Properties
    @property
    def owner(self) -> T:
        """Component Owner"""
        return self.__owner

    # endregion Properties
