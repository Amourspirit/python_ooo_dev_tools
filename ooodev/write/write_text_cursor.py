from __future__ import annotations
from typing import Any, cast, Sequence, overload, TYPE_CHECKING, TypeVar, Generic
import uno

from ooodev.mock import mock_g
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.paragraph_cursor_partial import ParagraphCursorPartial
from ooodev.adapter.text.sentence_cursor_partial import SentenceCursorPartial
from ooodev.adapter.style.character_properties_partial import CharacterPropertiesPartial
from ooodev.adapter.style.paragraph_properties_partial import ParagraphPropertiesPartial
from ooodev.adapter.text.text_cursor_comp import TextCursorComp
from ooodev.adapter.text.word_cursor_partial import WordCursorPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import write as mWrite
from ooodev.loader import lo as mLo
from ooodev.utils import selection as mSelection
from ooodev.utils.context.lo_context import LoContext
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.write.partial.text_cursor_partial import TextCursorPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.write.write_text import WriteText
from ooodev.proto.component_proto import ComponentT

if TYPE_CHECKING:
    from com.sun.star.text import XTextDocument
    from com.sun.star.text import XTextCursor
    from ooodev.proto.style_obj import StyleT
    from ooodev.write.style.direct.character_styler import CharacterStyler

T = TypeVar("T", bound="ComponentT")


class WriteTextCursor(
    LoInstPropsPartial,
    WriteDocPropPartial,
    EventsPartial,
    TextCursorPartial[T],
    Generic[T],
    TextCursorComp,
    ParagraphCursorPartial,
    SentenceCursorPartial,
    WordCursorPartial,
    CharacterPropertiesPartial,
    ParagraphPropertiesPartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
    PropPartial,
    QiPartial,
    TheDictionaryPartial,
    StylePartial,
):
    """
    Represents a writer text cursor.

    This class implements ``__len__()`` method, which returns the number of characters in the range.

    .. seealso::
        - :ref:`help_writer_format_direct_cursor_char_styler`
    """

    def __init__(self, owner: T, component: XTextCursor, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextCursor): A UNO object that supports ``com.sun.star.text.TextCursor`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        EventsPartial.__init__(self)
        TextCursorPartial.__init__(self, owner=self._owner, component=component)
        TextCursorComp.__init__(self, component)  # type: ignore
        ParagraphCursorPartial.__init__(self, component, None)  # type: ignore
        SentenceCursorPartial.__init__(self, component, None)  # type: ignore
        WordCursorPartial.__init__(self, component, None)  # type: ignore
        CharacterPropertiesPartial.__init__(self, component=component)  # type: ignore
        ParagraphPropertiesPartial.__init__(self, component=component)  # type: ignore
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        TheDictionaryPartial.__init__(self)
        StylePartial.__init__(self, component=component)
        self._style_direct_char = None

    def __len__(self) -> int:
        with LoContext(self.lo_inst):
            result = mSelection.Selection.range_len(cast("XTextDocument", self.owner.component), self.component)
        return result

    def get_write_text(self) -> WriteText[WriteTextCursor]:
        """
        Gets the WriteText (XText Component) object from the cursor.

        Returns:
            WriteText: WriteText object
        """

        return WriteText(self, self.get_text())

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
        return self._owner

    @property
    def style_direct_char(self) -> CharacterStyler:
        """
        Direct Character Styler.

        Returns:
            CharacterStyler: Character Styler
        """
        if self._style_direct_char is None:
            # pylint: disable=import-outside-toplevel
            from ooodev.write.style.direct.character_styler import CharacterStyler

            self._style_direct_char = CharacterStyler(write_doc=self.write_doc, component=self.component)
            self._style_direct_char.add_event_observers(self.event_observer)
        return self._style_direct_char

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.write.style.direct.character_styler import CharacterStyler
