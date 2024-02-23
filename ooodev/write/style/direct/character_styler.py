from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.style_named_event import StyleNameEvent
from ooodev.events.write_named_event import WriteNamedEvent
from ooodev.format.inner.partial.calc.font.font_effects_partial import FontEffectsPartial
from ooodev.format.inner.partial.font.font_only_partial import FontOnlyPartial
from ooodev.format.inner.partial.font.font_partial import FontPartial
from ooodev.format.inner.partial.font.font_position_partial import FontPositionPartial
from ooodev.format.inner.partial.write.char.borders.write_char_borders_partial import WriteCharBordersPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.format.writer.style import Char

if TYPE_CHECKING:
    from ...write_doc import WriteDoc


class CharacterStyler(
    WriteDocPropPartial,
    EventsPartial,
    FontOnlyPartial,
    FontEffectsPartial,
    FontPartial,
    WriteCharBordersPartial,
    FontPositionPartial,
    TheDictionaryPartial,
):
    def __init__(self, write_doc: WriteDoc, component: Any) -> None:
        WriteDocPropPartial.__init__(self, obj=write_doc)
        EventsPartial.__init__(self)

        FontOnlyPartial.__init__(
            self, factory_name="ooodev.write.char", component=component, lo_inst=self.write_doc.lo_inst
        )
        FontEffectsPartial.__init__(
            self, factory_name="ooodev.write.char", component=component, lo_inst=self.write_doc.lo_inst
        )
        FontPartial.__init__(
            self, factory_name="ooodev.write.char", component=component, lo_inst=self.write_doc.lo_inst
        )
        WriteCharBordersPartial.__init__(self, component=component)
        FontPositionPartial.__init__(
            self, factory_name="ooodev.write.char", component=component, lo_inst=self.write_doc.lo_inst
        )
        # The dictionary can be used to add extra data to the object. This is useful for event handling.
        TheDictionaryPartial.__init__(self)
        self._component = component
        self._init_events()

    def _init_events(self) -> None:
        self._fn_on_style_backup = self._on_style_backup
        self._fn_on_style_applying = self._on_style_applying
        self._fn_on_style_applied = self._on_style_applied
        self.subscribe_event(StyleNameEvent.STYLE_APPLIED, self._fn_on_style_applied)
        self.subscribe_event(StyleNameEvent.STYLE_APPLYING, self._fn_on_style_applying)
        self.subscribe_event("before_style_font_position_backup", self._fn_on_style_backup)

    def _on_style_backup(self, src: Any, event_args: CancelEventArgs) -> None:
        # by default event_args.cancel = True in this case
        event_args.cancel = False

    def _on_style_applying(self, src: Any, event_args: CancelEventArgs) -> None:
        args = CancelEventArgs(source=self)
        args.event_data = event_args.event_data.copy()
        self.trigger_event(WriteNamedEvent.CHARACTER_STYLE_APPLYING, args)
        event_args.cancel = args.cancel
        event_args.handled = args.handled
        event_args.event_data = args.event_data

    def _on_style_applied(self, src: Any, event_args: EventArgs) -> None:
        args = EventArgs(source=self)
        args.event_data = event_args.event_data.copy()
        self.trigger_event(WriteNamedEvent.CHARACTER_STYLE_APPLIED, args)
        event_args.event_data = args.event_data

    def clear(self) -> None:
        """Clears the formatting of the character."""
        # pylint: disable=no-member
        Char.default.apply(self._component)
