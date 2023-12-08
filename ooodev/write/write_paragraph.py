from __future__ import annotations
from typing import Any, cast, TypeVar, Generic, TYPE_CHECKING
import uno

from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_range_partial import TextRangePartial


from ooodev.adapter.text.paragraph_comp import ParagraphComp
from ooodev.adapter.text.text_content_comp import TextContentComp
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from . import write_text_portions as mWriteTextPortions

if TYPE_CHECKING:
    from com.sun.star.container import XEnumerationAccess

T = TypeVar("T", bound="ComponentT")


class WriteParagraph(
    Generic[T],
    TextContentComp,
    ParagraphComp,
    PropertyChangeImplement,
    VetoableChangeImplement,
    TextRangePartial,
    QiPartial,
    PropPartial,
):
    """
    Represents writer paragraph content.

    Contains Enumeration Access.
    """

    def __init__(self, owner: T, component: Any) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (Any): UNO object that supports ``com.sun.star.text.Paragraph`` service.
        """
        self.__owner = owner
        TextContentComp.__init__(self, component)
        ParagraphComp.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)  # type: ignore
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)  # type: ignore
        TextRangePartial.__init__(self, component=self.component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    def get_text_portions(self) -> mWriteTextPortions.WriteTextPortions[WriteParagraph[T]]:
        """Returns the text portions of this paragraph."""
        return mWriteTextPortions.WriteTextPortions(owner=self, component=cast("XEnumerationAccess", self.component))

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
