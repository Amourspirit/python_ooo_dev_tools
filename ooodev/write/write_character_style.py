from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


if TYPE_CHECKING:
    from com.sun.star.style import XStyle
    from ooodev.proto.component_proto import ComponentT

    T = TypeVar("T", bound="ComponentT")

from ooodev.adapter.style.character_style_comp import CharacterStyleComp
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils import lo as mLo


class WriteCharacterStyle(Generic[T], CharacterStyleComp, QiPartial, PropPartial):
    """Represents writer Character Style."""

    def __init__(self, owner: T, component: XStyle) -> None:
        """
        Constructor

        Args:
            owner (WriteDoc): Doc that owns this component.
            component (XStyle): UNO object that supports ``com.sun.star.style.CharacterStyle`` service.
        """
        self.__owner = owner
        CharacterStyleComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    # region Properties
    @property
    def owner(self) -> T:
        """Doc that owns this component."""
        return self.__owner

    # endregion Properties
