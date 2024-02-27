from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno

from ooodev.loader.inst.lo_inst import LoInst
from ooodev.adapter.style.style_comp import StyleComp
from ooodev.loader import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.proto.component_proto import ComponentT

if TYPE_CHECKING:
    from com.sun.star.style import XStyle

T = TypeVar("T", bound="ComponentT")


class WriteStyle(Generic[T], StyleComp, QiPartial, PropPartial):
    """Represents writer Style."""

    def __init__(self, owner: T, component: XStyle, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XStyle): UNO object that supports ``com.sun.star.style.Style`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            self._lo_inst = mLo.Lo.current_lo
        else:
            self._lo_inst = lo_inst
        self._owner = owner
        StyleComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self._lo_inst)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=self._lo_inst)  # type: ignore
        # self.__doc = doc

    # region Properties
    @property
    def owner(self) -> T:
        """Component Owner"""
        return self._owner

    # endregion Properties
