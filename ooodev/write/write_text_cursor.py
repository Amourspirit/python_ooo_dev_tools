from __future__ import annotations
from typing import Any, cast, Sequence, overload, TYPE_CHECKING, TypeVar, Generic
import uno

if TYPE_CHECKING:
    from com.sun.star.text import XTextDocument
    from com.sun.star.text import XTextCursor
    from ooodev.proto.style_obj import StyleT
    from ooodev.units import UnitT
    from ooodev.utils.type_var import PathOrStr, Table

from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_cursor_comp import TextCursorComp
from ooodev.office import write as mWrite
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils import selection as mSelection
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from .partial.text_cursor_partial import TextCursorPartial


T = TypeVar("T", bound="ComponentT")


class WriteTextCursor(
    Generic[T],
    TextCursorPartial,
    TextCursorComp,
    PropertyChangeImplement,
    VetoableChangeImplement,
    PropPartial,
    QiPartial,
):
    """
    Represents a writer text cursor.

    This class implements ``__len__()`` method, which returns the number of characters in the range.
    """

    def __init__(self, owner: T, component: XTextCursor) -> None:
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

    # region style_prev_paragraph()
    @overload
    def style_prev_paragraph(self, styles: Sequence[StyleT]) -> None:
        """
        Style previous paragraph.

        Args:
            styles (Sequence[StyleT]): One or more styles to apply to text.

        Returns:
            None:
        """
        ...

    @overload
    def style_prev_paragraph(self, prop_val: Any) -> None:
        """
        Style previous paragraph.

        Args:
            prop_val (Any): Property value. Applied to ``ParaStyleName`` property.

        Returns:
            None:
        """
        ...

    @overload
    def style_prev_paragraph(self, prop_val: Any, prop_name: str) -> None:
        """
        Style previous paragraph.

        Args:
            prop_val (Any): Property value.
            prop_name (str): Property Name. This is the property name to apply the style to.

        Returns:
            None:
        """
        ...

    def style_prev_paragraph(self, *args, **kwargs) -> None:
        """
        Style previous paragraph.

        Args:
            styles (Sequence[StyleT]): One or more styles to apply to text.
            prop_val (Any): Property value.
            prop_name (str): Property Name.

        :events:
            If using styles then the following events are triggered for each style.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-event`

            Otherwise the following events are triggered once.

            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLING` :eventref:`src-docs-key-event-cancel`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.STYLED` :eventref:`src-docs-key-event`

        Returns:
            None:

        .. collapse:: Example

            .. code-block:: python

                doc.style_prev_paragraph(prop_val=ParagraphAdjust.CENTER, prop_name="ParaAdjust")
        """
        mWrite.Write.style_prev_paragraph(self.component, *args, **kwargs)

    # endregion style_prev_paragraph()

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
