from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


from ooodev.adapter.style.paragraph_style_comp import ParagraphStyleComp
from ooodev.proto.component_proto import ComponentT
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial

if TYPE_CHECKING:
    from com.sun.star.style import XStyle

T = TypeVar("T", bound="ComponentT")


class WriteParagraphStyle(Generic[T], ParagraphStyleComp, QiPartial, PropPartial):
    """Represents writer Paragraph Style."""

    def __init__(self, owner: T, component: XStyle, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XStyle): UNO object that supports ``com.sun.star.style.ParagraphStyle`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        self._lo_inst = mLo.Lo.current_lo if lo_inst is None else lo_inst
        self._owner = owner
        ParagraphStyleComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self._lo_inst)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=self._lo_inst)  # type: ignore
        # self.__doc = doc

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
