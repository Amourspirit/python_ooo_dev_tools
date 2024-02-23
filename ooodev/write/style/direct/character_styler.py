from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.format.inner.partial.calc.font.font_effects_partial import FontEffectsPartial
from ooodev.format.inner.partial.font.font_only_partial import FontOnlyPartial
from ooodev.format.inner.partial.font.font_partial import FontPartial
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.partial.write.char.borders.write_char_borders_partial import WriteCharBordersPartial
from ooodev.format.inner.partial.font.font_position_partial import FontPositionPartial

if TYPE_CHECKING:
    from ...write_doc import WriteDoc
    from ooodev.events.args.cancel_event_args import CancelEventArgs


class CharacterStyler(
    WriteDocPropPartial,
    EventsPartial,
    FontOnlyPartial,
    FontEffectsPartial,
    FontPartial,
    WriteCharBordersPartial,
    FontPositionPartial,
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
        self._init_events()

    def _init_events(self) -> None:
        self._fn_on_style_backup = self._on_style_backup
        self.subscribe_event("before_style_font_position_backup", self._fn_on_style_backup)

    def _on_style_backup(self, src: Any, event_args: CancelEventArgs) -> None:
        # by default event_args.cancel = True in this case
        event_args.cancel = False
