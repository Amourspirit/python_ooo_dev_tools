from __future__ import annotations
from typing import Any, TypeVar, Generic, TYPE_CHECKING
import uno

from ooodev.adapter.text.text_portion_comp import TextPortionComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial

if TYPE_CHECKING:
    from ooodev.proto.component_proto import ComponentT

T = TypeVar("T", bound="ComponentT")


class WriteTextPortion(
    Generic[T],
    LoInstPropsPartial,
    TextPortionComp,
    WriteDocPropPartial,
    QiPartial,
    ServicePartial,
    TheDictionaryPartial,
    PropPartial,
    StylePartial,
):
    """
    Represents writer paragraph content.

    Contains Enumeration Access.
    """

    def __init__(self, owner: T, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (Any): UNO object that supports ``com.sun.star.text.TextPortion`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        TextPortionComp.__init__(self, component)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        TheDictionaryPartial.__init__(self)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)
        # self.__doc = doc

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
