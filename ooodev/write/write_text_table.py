from __future__ import annotations
from typing import cast, TYPE_CHECKING, TypeVar, Generic
import uno


if TYPE_CHECKING:
    from com.sun.star.text import XTextContent
    from com.sun.star.container import XEnumerationAccess

from ooodev.adapter.text.text_table_comp import TextTableComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils.partial.qi_partial import QiPartial
from . import write_text_portions as mWriteTextPortions

T = TypeVar("T", bound="ComponentT")


class WriteTextTable(Generic[T], TextTableComp, QiPartial, StylePartial):
    """Represents writer text content."""

    def __init__(self, owner: T, component: XTextContent, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextContent): UNO object that supports ``com.sun.star.text.TextContent`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            self._lo_inst = mLo.Lo.current_lo
        else:
            self._lo_inst = lo_inst
        self._owner = owner
        TextTableComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self._lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)

    def get_text_portions(self) -> mWriteTextPortions.WriteTextPortions[T]:
        """Returns the text portions of this paragraph."""
        return mWriteTextPortions.WriteTextPortions(
            owner=self.owner, component=cast("XEnumerationAccess", self.component), lo_inst=self._lo_inst
        )

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
