from __future__ import annotations
from typing import cast, TYPE_CHECKING, TypeVar, Generic
import uno

from com.sun.star.text import XTextViewCursor

if TYPE_CHECKING:
    from com.sun.star.text import XTextViewCursor
    from com.sun.star.text import XTextDocument

from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_view_cursor_comp import TextViewCursorComp
from ooodev.adapter.view.line_cursor_partial import LineCursorPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import write as mWrite
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils import selection as mSelection
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial

from .partial.text_cursor_partial import TextCursorPartial

T = TypeVar("T", bound="ComponentT")


class WriteTextViewCursor(
    Generic[T],
    TextCursorPartial,
    TextViewCursorComp,
    LineCursorPartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
    PropPartial,
    QiPartial,
    StylePartial,
):
    """Represents a writer text view cursor."""

    def __init__(self, owner: T, component: XTextViewCursor) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextViewCursor): A UNO object that supports ``com.sun.star.text.TextViewCursor`` service.
        """
        self.__owner = owner
        TextCursorPartial.__init__(self, owner=owner, component=component)
        TextViewCursorComp.__init__(self, component)  # type: ignore
        LineCursorPartial.__init__(self, component, None)  # type: ignore
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        StylePartial.__init__(self, component=component)

    def __len__(self) -> int:
        return mSelection.Selection.range_len(cast("XTextDocument", self.owner.component), self.component)

    def get_coord_str(self) -> str:
        """
        Gets coordinates for cursor in format such as ``"10, 10"``

        Returns:
            str: coordinates as string
        """
        return mWrite.Write.get_coord_str(self.component)  # type: ignore

    def get_current_page_num(self) -> int:
        """
        Gets current page number.

        Returns:
            int: current page number
        """
        return mWrite.Write.get_current_page(self.component)  # type: ignore

    def get_text_view_cursor(self) -> XTextViewCursor:
        """
        Gets the text view cursor.

        Returns:
            XTextViewCursor: text view cursor
        """
        return self.qi(XTextViewCursor, True)

    def get_current_page(self) -> int:
        """
        Gets the current page


        Returns:
            int: Page number if present; Otherwise, -1

        See Also:
            :py:meth:`~.Write.get_page_number`
        """
        return mWrite.Write.get_current_page(self.get_text_view_cursor())

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
