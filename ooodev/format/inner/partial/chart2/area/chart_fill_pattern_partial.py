from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.format.inner.style_factory import chart2_area_pattern_factory
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
    from com.sun.star.awt import XBitmap
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.chart2.area.chart_fill_pattern_t import ChartFillPatternT
    from ooodev.format.inner.preset.preset_pattern import PresetPatternKind
else:
    XChartDocument = Any
    XBitmap = Any
    ChartFillPatternT = Any
    PresetPatternKind = Any


class ChartFillPatternPartial:
    """
    Partial class for Chart Fill Pattern.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_area_pattern",
            after_event="after_style_area_pattern",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def _ChartFillPatternPartial__get_chart_doc(self) -> XChartDocument:
        if isinstance(self, ChartDocPropPartial):
            return self.chart_doc.component

        raise NotImplementedError

    def style_area_pattern(
        self,
        *,
        bitmap: XBitmap | None = None,
        name: str = "",
        tile: bool = True,
        stretch: bool = False,
        auto_name: bool = False,
    ) -> ChartFillPatternT | None:
        """
        Style Area Color.

        Args:
            bitmap (XBitmap, optional): Bitmap instance. If ``name`` is not already in the Bitmap Table then this property is required.
            name (str, optional): Specifies the name of the pattern. This is also the name that is used to store bitmap in LibreOffice Bitmap Table.
            tile (bool, optional): Specified if bitmap is tiled. Defaults to ``True``.
            stretch (bool, optional): Specifies if bitmap is stretched. Defaults to ``False``.
            auto_name (bool, optional): Specifies if ``name`` is ensured to be unique. Defaults to ``False``.

        Raises:
            CancelEventError: If the event ``before_style_area_pattern`` is cancelled and not handled.

        Returns:
            ChartFillPatternT | None: Fill Image instance or ``None`` if cancelled.
        """
        doc = self._ChartFillPatternPartial__get_chart_doc()
        factory = chart2_area_pattern_factory
        kwargs = {
            "chart_doc": doc,
            "name": name,
            "tile": tile,
            "stretch": stretch,
            "auto_name": auto_name,
        }
        if bitmap is not None:
            kwargs["bitmap"] = bitmap
        return self.__styler.style(factory=factory, **kwargs)

    def style_area_pattern_get(self) -> ChartFillPatternT | None:
        """
        Gets the Area Area Pattern Style.

        Raises:
            CancelEventError: If the event ``before_style_area_pattern_get`` is cancelled and not handled.

        Returns:
            ChartFillPatternT | None: Area pattern style or ``None`` if cancelled.
        """
        doc = self._ChartFillPatternPartial__get_chart_doc()
        return self.__styler.style_get(factory=chart2_area_pattern_factory, chart_doc=doc)

    def style_area_pattern_from_preset(self, preset: PresetPatternKind) -> ChartFillPatternT | None:
        """
        Style Area Gradient from Preset.

        Args:
            preset (PresetPatternKind): Preset Image Kind.

        Returns:
            ChartFillPatternT: Chart Fill Image instance.

        Hint:
            - ``PresetPatternKind`` can be imported from ``ooodev.format.inner.preset.preset_pattern``
        """
        styler = self.__styler
        doc = self._ChartFillPatternPartial__get_chart_doc()
        fe = styler.style_get(
            factory=chart2_area_pattern_factory,
            call_method_name="from_preset",
            event_name_suffix="_from_preset",
            obj_arg_name="",
            chart_doc=doc,
            preset=preset,
        )
        styler.style_apply(style=fe, chart_doc=doc, preset=preset)
        return fe
