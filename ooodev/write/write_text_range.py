from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno


if TYPE_CHECKING:
    from com.sun.star.text import XTextRange

from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_range_comp import TextRangeComp
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial

T = TypeVar("T", bound="ComponentT")


class WriteTextRange(
    Generic[T], TextRangeComp, PropertyChangeImplement, VetoableChangeImplement, QiPartial, PropPartial
):
    """Represents writer TextRange."""

    def __init__(self, owner: T, component: XTextRange) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextRange): UNO object that supports ``com.sun.star.text.TextRange`` service.
        """
        self.__owner = owner
        TextRangeComp.__init__(self, component)  # type: ignore
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
