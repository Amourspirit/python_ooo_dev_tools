from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


if TYPE_CHECKING:
    from com.sun.star.style import XStyle

from ooodev.adapter.style.character_style_comp import CharacterStyleComp
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.inst.lo.lo_inst import LoInst

T = TypeVar("T", bound="ComponentT")


class WriteCharacterStyle(Generic[T], CharacterStyleComp, QiPartial, PropPartial):
    """Represents writer Character Style."""

    def __init__(self, owner: T, component: XStyle, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XStyle): UNO object that supports ``com.sun.star.style.CharacterStyle`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            self._lo_inst = mLo.Lo.current_lo
        else:
            self._lo_inst = lo_inst
        self._owner = owner
        CharacterStyleComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self._lo_inst)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=self._lo_inst)  # type: ignore

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
