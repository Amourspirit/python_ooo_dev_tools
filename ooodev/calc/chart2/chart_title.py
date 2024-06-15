from __future__ import annotations
from typing import Any, TYPE_CHECKING, TypeVar, Generic
import contextlib
import uno

from ooodev.mock import mock_g
from ooodev.adapter.chart2.title_comp import TitleComp
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.partial.font.font_effects_partial import FontEffectsPartial
from ooodev.format.inner.partial.font.font_only_partial import FontOnlyPartial
from ooodev.format.inner.partial.font.font_partial import FontPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.office import chart2 as mChart2
from ooodev.proto.component_proto import ComponentT
from ooodev.units.angle import Angle
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.format.inner.partial.chart2.title.alignment.chart2_title_orientation_partial import (
    Chart2TitleOrientationPartial,
)
from ooodev.format.inner.partial.area.fill_color_partial import FillColorPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_gradient_partial import ChartFillGradientPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_hatch_partial import ChartFillHatchPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_img_partial import ChartFillImgPartial
from ooodev.format.inner.partial.chart2.area.chart_fill_pattern_partial import ChartFillPatternPartial
from ooodev.format.inner.partial.chart2.borders.border_line_properties_partial import BorderLinePropertiesPartial

from ooodev.format.inner.partial.position_size.chart2.chart2_position_partial import Chart2PositionPartial
from ooodev.calc.chart2.kind.chart_title_kind import ChartTitleKind
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial

# from ooodev.format.inner.partial.position_size.draw.position_partial import PositionPartial

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.proto.style_obj import StyleT
    from ooodev.events.args.cancel_event_args import CancelEventArgs
    from ooodev.calc.chart2.chart_doc import ChartDoc

_T = TypeVar("_T", bound="ComponentT")


class ChartTitle(
    Generic[_T],
    LoInstPropsPartial,
    EventsPartial,
    TitleComp,
    ChartDocPropPartial,
    PropPartial,
    QiPartial,
    ServicePartial,
    TheDictionaryPartial,
    CalcDocPropPartial,
    CalcSheetPropPartial,
    StylePartial,
    FontOnlyPartial,
    FontEffectsPartial,
    FontPartial,
    Chart2TitleOrientationPartial,
    FillColorPartial,
    ChartFillGradientPartial,
    ChartFillHatchPartial,
    ChartFillImgPartial,
    ChartFillPatternPartial,
    BorderLinePropertiesPartial,
    Chart2PositionPartial,
):
    """
    Class for managing Chart2 Chart Title Component.
    """

    def __init__(
        self,
        owner: _T,
        chart_doc: ChartDoc,
        component: Any,
        title_kind: ChartTitleKind,
        lo_inst: LoInst | None = None,
    ) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Chart2 Title Component.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        EventsPartial.__init__(self)
        TitleComp.__init__(self, component=component)
        ChartDocPropPartial.__init__(self, chart_doc=chart_doc)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        TheDictionaryPartial.__init__(self)
        CalcDocPropPartial.__init__(self, obj=chart_doc.calc_doc)
        CalcSheetPropPartial.__init__(self, obj=chart_doc.calc_sheet)
        StylePartial.__init__(self, component=component)
        FontEffectsPartial.__init__(self, factory_name="ooodev.chart2.title", component=component, lo_inst=lo_inst)
        FontOnlyPartial.__init__(self, factory_name="ooodev.chart2.title", component=component, lo_inst=lo_inst)
        FontPartial.__init__(self, factory_name="ooodev.general_style.text", component=component, lo_inst=lo_inst)
        Chart2TitleOrientationPartial.__init__(self, component=component)
        FillColorPartial.__init__(self, factory_name="ooodev.char2.title.area", component=component, lo_inst=lo_inst)
        ChartFillGradientPartial.__init__(
            self, factory_name="ooodev.char2.title", component=component, lo_inst=lo_inst
        )
        ChartFillHatchPartial.__init__(self, factory_name="ooodev.char2.title", component=component, lo_inst=lo_inst)
        ChartFillPatternPartial.__init__(self, factory_name="ooodev.char2.title", component=component, lo_inst=lo_inst)
        ChartFillImgPartial.__init__(self, factory_name="ooodev.char2.title", component=component, lo_inst=lo_inst)
        BorderLinePropertiesPartial.__init__(
            self, factory_name="ooodev.char2.title", component=component, lo_inst=lo_inst
        )
        Chart2PositionPartial.__init__(self, factory_name="ooodev.chart2.title", component=component, lo_inst=lo_inst)
        # PositionPartial.__init__(self, factory_name="ooodev.draw.position", component=component, lo_inst=lo_inst)
        self._owner = owner
        self._title_kind = title_kind
        self._init_events()

    # region Events
    def _init_events(self) -> None:
        self._fn_on_apply_style_text = self._on_apply_style_text
        self._fn_on_global_cancel = self._on_global_cancel
        self._fn_on_before_style_position = self._on_before_style_position
        self.subscribe_event("before_style_font_effect", self._fn_on_apply_style_text)
        self.subscribe_event("before_style_font_only", self._fn_on_apply_style_text)
        self.subscribe_event("before_style_general_font", self._fn_on_apply_style_text)
        self.subscribe_event(GblNamedEvent.EVENT_CANCELED, self._fn_on_global_cancel)
        self.subscribe_event("before_style_position", self._fn_on_before_style_position)

    def _on_apply_style_text(self, source: Any, args: CancelEventArgs) -> None:
        fo_strs = self.component.getText()
        if fo_strs:
            args.event_data["this_component"] = fo_strs[0]
        else:
            args.cancel = True

    def _on_global_cancel(self, source: Any, args: CancelEventArgs) -> None:
        initial_event = args.get("initial_event", "")

        if initial_event in {"before_style_font_effect", "before_style_font_only", "before_style_general_font"}:
            args.handled = True

    def _on_before_style_position(self, source: Any, args: CancelEventArgs) -> None:
        # get the old chart instance from the chart document.
        from com.sun.star.chart import XChartDocument

        if self.title_kind == ChartTitleKind.TITLE:
            with contextlib.suppress(Exception):
                # get the shape from the chart document
                doc = self.lo_inst.qi(XChartDocument, self.chart_doc.component, True)

                shape = doc.getTitle()  # shape
                args.event_data["this_component"] = shape
        elif self.title_kind == ChartTitleKind.SUBTITLE:
            with contextlib.suppress(Exception):
                # get the shape from the chart document
                doc = self.lo_inst.qi(XChartDocument, self.chart_doc.component, True)
                shape = doc.getSubTitle()  # shape
                args.event_data["this_component"] = shape
        else:
            # cancel the event, No Shape to work with.
            args.cancel = True
            args.handled = True

    # endregion Events

    # region StylePartial Overrides

    def apply_styles(self, *styles: StyleT, **kwargs) -> None:
        """
        Applies style to component.

        Args:
            styles expandable list of styles object such as ``Font`` to apply to ``obj``.
            kwargs (Any, optional): Expandable list of key value pairs.

        Returns:
            None:
        """
        mChart2.Chart2._style_title(self.component, styles)

    # endregion

    @property
    def owner(self) -> _T:
        """Chart Document"""
        return self._owner

    @property
    def rotation(self) -> Angle:
        """Gets or sets the rotation angle of the title."""
        return Angle(self.get_property("TextRotation", 0))

    @rotation.setter
    def rotation(self, value: Angle | int) -> None:
        rotation = Angle(int(value))
        self.set_property(TextRotation=rotation.value)

    @property
    def title_kind(self) -> ChartTitleKind:
        """Gets the title kind."""
        return self._title_kind


if mock_g.FULL_IMPORT:
    from com.sun.star.chart import XChartDocument
