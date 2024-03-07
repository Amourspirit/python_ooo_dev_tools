from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.style_named_event import StyleNameEvent
from ooodev.format.inner.partial.area.fill_color_partial import FillColorPartial
from ooodev.format.inner.partial.write.area.write_table_fill_img_partial import WriteTableFillImgPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial
from ooodev.events.write_named_event import WriteNamedEvent
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial


if TYPE_CHECKING:
    from ooodev.write.table.write_table import WriteTable


class RowStyler(
    WriteDocPropPartial,
    WriteTablePropPartial,
    EventsPartial,
    LoInstPropsPartial,
    FillColorPartial,
    WriteTableFillImgPartial,
    TheDictionaryPartial,
):
    """
    Row Styler class.

    Class set various Row properties.

    Events are raises when a style is being applied and when a style has been applied.
    The ``WriteNamedEvent.TABLE_STYLE_APPLYING`` event is raised before a style is applied.
    The ``WriteNamedEvent.TABLE_STYLE_APPLIED`` event is raised after a style has been applied.
    The event data is a dictionary that contains the following:

    - ``cancel_apply``: If set to ``True`` the style will not be applied. This is only used in the ``TABLE_STYLE_APPLYING`` event.
    - ``this_component``: The component that the style is being applied to. This is the normally same component passed to the constructor.
    - ``styler_object``: The style that is being applied. This is ``None`` when ``TABLE_STYLE_APPLYING`` is raised. Is the style that was applied when ``TABLE_STYLE_APPLIED`` is raised.

    Other style specific data is also in the dictionary such as the parameter values used to apply the style.

    The ``event_args.event_source`` is the instance of the ``RowStyler`` class.
    """

    # pylint: disable=unused-argument

    def __init__(self, owner: WriteTable, component: Any) -> None:
        """
        Constructor.

        Args:
            owner (WriteTable): Write Table instance.
            component (Any): component instance.
        """
        WriteTablePropPartial.__init__(self, obj=owner)
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)
        LoInstPropsPartial.__init__(self, lo_inst=owner.write_doc.lo_inst)
        EventsPartial.__init__(self)

        FillColorPartial.__init__(
            self, factory_name="ooodev.write.table.background", component=component, lo_inst=self.write_doc.lo_inst
        )
        WriteTableFillImgPartial.__init__(
            self, factory_name="ooodev.write.table.background", component=component, lo_inst=self.write_doc.lo_inst
        )

        # The dictionary can be used to add extra data to the object. This is useful for event handling.
        TheDictionaryPartial.__init__(self)
        self._component = component
        self._init_events()

    def _init_events(self) -> None:
        self._fn_on_style_applying = self._on_style_applying
        self._fn_on_style_applied = self._on_style_applied
        self.subscribe_event(StyleNameEvent.STYLE_APPLIED, self._fn_on_style_applied)
        self.subscribe_event(StyleNameEvent.STYLE_APPLYING, self._fn_on_style_applying)

    def _on_style_applying(self, src: Any, event_args: CancelEventArgs) -> None:
        args = CancelEventArgs(source=self)
        args.event_data = event_args.event_data.copy()
        self.trigger_event(WriteNamedEvent.TABLE_STYLE_APPLYING, args)
        event_args.cancel = args.cancel
        event_args.handled = args.handled
        event_args.event_data = args.event_data

    def _on_style_applied(self, src: Any, event_args: EventArgs) -> None:
        args = EventArgs(source=self)
        args.event_data = event_args.event_data.copy()
        self.trigger_event(WriteNamedEvent.TABLE_STYLE_APPLIED, args)
        event_args.event_data = args.event_data
