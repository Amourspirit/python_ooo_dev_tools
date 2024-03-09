from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.text import XSimpleText

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.events.style_named_event import StyleNameEvent
from ooodev.format.inner.partial.area.fill_color_partial import FillColorPartial
from ooodev.format.inner.partial.write.area.write_table_fill_img_partial import WriteTableFillImgPartial
from ooodev.format.inner.partial.write.table.write_table_cell_borders_partial import WriteTableCellBordersPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial
from ooodev.events.write_named_event import WriteNamedEvent
from ooodev.format.inner.partial.font.font_only_partial import FontOnlyPartial
from ooodev.format.inner.partial.font.font_partial import FontPartial
from ooodev.format.inner.partial.write.para.write_para_alignment_partial import WriteParaAlignmentPartial
from ooodev.format.inner.partial.calc.font.font_effects_partial import FontEffectsPartial

# from ooodev.format.inner.partial.numbers.numbers_numbers_partial import NumbersNumbersPartial
from ooodev.format.inner.partial.write.numbers.numbers_numbers_partial import NumbersNumbersPartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.utils.partial.qi_partial import QiPartial


if TYPE_CHECKING:
    from ooodev.write.table.write_table import WriteTable


class CellStyler(
    WriteDocPropPartial,
    WriteTablePropPartial,
    EventsPartial,
    LoInstPropsPartial,
    QiPartial,
    FillColorPartial,
    WriteTableFillImgPartial,
    WriteTableCellBordersPartial,
    FontOnlyPartial,
    FontEffectsPartial,
    FontPartial,
    NumbersNumbersPartial,
    WriteParaAlignmentPartial,
    TheDictionaryPartial,
):
    """
    Cell Styler class.

    Class set various cell properties.

    Events are raises when a style is being applied and when a style has been applied.
    The ``WriteNamedEvent.TABLE_STYLE_APPLYING`` event is raised before a style is applied.
    The ``WriteNamedEvent.TABLE_STYLE_APPLIED`` event is raised after a style has been applied.
    The event data is a dictionary that contains the following:

    - ``cancel_apply``: If set to ``True`` the style will not be applied. This is only used in the ``TABLE_STYLE_APPLYING`` event.
    - ``this_component``: The component that the style is being applied to. This is the normally same component passed to the constructor.
    - ``styler_object``: The style that is being applied. This is ``None`` when ``TABLE_STYLE_APPLYING`` is raised. Is the style that was applied when ``TABLE_STYLE_APPLIED`` is raised.

    Other style specific data is also in the dictionary such as the parameter values used to apply the style.

    The ``event_args.event_source`` is the instance of the ``CellStyler`` class.
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
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)

        FillColorPartial.__init__(
            self, factory_name="ooodev.write.table.background", component=component, lo_inst=self.lo_inst
        )
        WriteTableFillImgPartial.__init__(
            self, factory_name="ooodev.write.table.background", component=component, lo_inst=self.lo_inst
        )
        WriteTableCellBordersPartial.__init__(self, component=component)
        FontOnlyPartial.__init__(self, factory_name="ooodev.calc.cell", component=component, lo_inst=self.lo_inst)
        FontEffectsPartial.__init__(self, factory_name="ooodev.calc.cell", component=component, lo_inst=self.lo_inst)
        FontPartial.__init__(self, factory_name="ooodev.general_style.text", component=component, lo_inst=self.lo_inst)
        NumbersNumbersPartial.__init__(
            self, factory_name="ooodev.number.numbers", component=component, lo_inst=self.lo_inst
        )
        WriteParaAlignmentPartial.__init__(self, component=component)

        # The dictionary can be used to add extra data to the object. This is useful for event handling.
        TheDictionaryPartial.__init__(self)
        self._component = component
        self._init_events()

    def _init_events(self) -> None:
        self._fn_on_style_applying = self._on_style_applying
        self._fn_on_style_applied = self._on_style_applied
        self._fn_on_cursor_styling = self._on_cursor_styling
        self._fn_on_cursor_styled = self._on_cursor_styled
        self.subscribe_event(StyleNameEvent.STYLE_APPLIED, self._fn_on_style_applied)
        self.subscribe_event(StyleNameEvent.STYLE_APPLYING, self._fn_on_style_applying)
        self.subscribe_event("before_style_general_font", self._fn_on_cursor_styling)
        self.subscribe_event("after_style_general_font", self._fn_on_cursor_styled)
        self.subscribe_event("before_style_font_effect", self._fn_on_cursor_styling)
        self.subscribe_event("after_style_font_effect", self._fn_on_cursor_styled)
        self.subscribe_event("before_style_font_only", self._fn_on_cursor_styling)
        self.subscribe_event("after_style_font_only", self._fn_on_cursor_styled)
        self.subscribe_event("before_paragraph_alignment", self._fn_on_cursor_styling)
        self.subscribe_event("after_paragraph_alignment", self._fn_on_cursor_styled)

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

    def _on_cursor_styling(self, src: Any, event_args: CancelEventArgs) -> None:
        # com.sun.star.text.CellProperties
        comp = self.qi(XSimpleText, True)
        cursor = comp.createTextCursor()
        cursor.gotoStart(False)
        cursor.gotoEnd(True)
        event_args.event_data["this_component"] = cursor

    def _on_cursor_styled(self, src: Any, event_args: EventArgs) -> None:
        pass
