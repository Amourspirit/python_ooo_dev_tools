from __future__ import annotations
from typing import cast, TYPE_CHECKING, TypeVar, Generic
import uno

from ooodev.utils.context.lo_context import LoContext
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.drawing.shape_partial_props import ShapePartialProps
from ooodev.adapter.text.text_cursor_comp import TextCursorComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.utils import selection as mSelection
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.write.partial.text_cursor_partial import TextCursorPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.proto.component_proto import ComponentT

if TYPE_CHECKING:
    from com.sun.star.text import XTextDocument
    from com.sun.star.text import XTextCursor
    from ooodev.loader.inst.lo_inst import LoInst

_T = TypeVar("_T", bound="ComponentT")


class ShapeTextCursor(
    LoInstPropsPartial,
    TextCursorPartial[_T],
    Generic[_T],
    TextCursorComp,
    ShapePartialProps,
    PropertyChangeImplement,
    VetoableChangeImplement,
    PropPartial,
    QiPartial,
    StylePartial,
):
    """
    Represents a text cursor.

    This class implements ``__len__()`` method, which returns the number of characters in the range.
    """

    def __init__(self, owner: _T, component: XTextCursor, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextCursor): A UNO object that supports ``com.sun.star.text.TextCursor`` service.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo

        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        TextCursorPartial.__init__(self, owner=owner, component=component)
        TextCursorComp.__init__(self, component)  # type: ignore
        ShapePartialProps.__init__(self, component=component)  # type: ignore
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)

    def __len__(self) -> int:
        with LoContext(self.lo_inst):
            result = mSelection.Selection.range_len(cast("XTextDocument", self.owner.component), self.component)
        return result

    # region Properties
    @property
    def owner(self) -> _T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
