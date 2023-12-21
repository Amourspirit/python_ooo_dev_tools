from __future__ import annotations
from typing import cast, TYPE_CHECKING, TypeVar, Generic
import uno

if TYPE_CHECKING:
    from com.sun.star.text import XTextDocument
    from com.sun.star.text import XTextCursor

from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_cursor_comp import TextCursorComp
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils import selection as mSelection
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.text_cursor_partial import TextCursorPartial


_T = TypeVar("_T", bound="ComponentT")


class DrawTextCursor(
    Generic[_T],
    TextCursorPartial,
    TextCursorComp,
    PropertyChangeImplement,
    VetoableChangeImplement,
    PropPartial,
    QiPartial,
):
    """
    Represents a text cursor.

    This class implements ``__len__()`` method, which returns the number of characters in the range.
    """

    def __init__(self, owner: _T, component: XTextCursor) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextCursor): A UNO object that supports ``com.sun.star.text.TextCursor`` service.
        """
        self.__owner = owner
        TextCursorPartial.__init__(self, owner=owner, component=component)
        TextCursorComp.__init__(self, component)  # type: ignore
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        # self.__doc = doc

    def __len__(self) -> int:
        return mSelection.Selection.range_len(cast("XTextDocument", self.owner.component), self.component)

    # region Properties
    @property
    def owner(self) -> _T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
