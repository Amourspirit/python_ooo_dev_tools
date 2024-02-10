from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_factory import chart2_area_pattern_factory
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
    from com.sun.star.awt import XBitmap
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.area.chart2.chart_fill_pattern_t import ChartFillPatternT
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
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

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
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_area_pattern.__qualname__)
            event_data: Dict[str, Any] = {
                "bitmap": bitmap,
                "name": name,
                "tile": tile,
                "stretch": stretch,
                "auto_name": auto_name,
                "factory_name": factory_name,
                "component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_area_pattern", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_area_pattern")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style Area Image has been cancelled.")
                    else:
                        return None
                else:
                    return None
            bitmap = cargs.event_data.get("bitmap", bitmap)
            name = cargs.event_data.get("name", name)
            tile = cargs.event_data.get("tile", tile)
            stretch = cargs.event_data.get("stretch", stretch)
            auto_name = cargs.event_data.get("auto_name", auto_name)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("component", comp)

        styler = chart2_area_pattern_factory(factory_name)
        fe = styler(
            chart_doc=doc,
            bitmap=bitmap,
            name=name,
            tile=tile,
            stretch=stretch,
            auto_name=auto_name,
        )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_area_pattern", EventArgs.from_args(cargs))  # type: ignore
        return fe

    def style_area_pattern_from_preset(self, preset: PresetPatternKind) -> ChartFillPatternT | None:
        """
        Style Area Gradient from Preset.

        Args:
            preset (PresetPatternKind): Preset Image Kind.

        Returns:
            ChartFillPatternT: Chart Fill Image instance.
        """
        doc = self._ChartFillPatternPartial__get_chart_doc()
        cargs = None
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        comp = self.__component
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_area_pattern_from_preset.__qualname__)
            event_data: Dict[str, Any] = {
                "preset": preset,
                "factory_name": factory_name,
                "component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_area_pattern_from_preset", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_area_pattern_from_preset")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style Area Image has been cancelled.")
                    else:
                        return None
                else:
                    return None
            preset = cargs.event_data.get("preset", preset)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("component", comp)

        styler = chart2_area_pattern_factory(factory_name)
        fe = styler.from_preset(chart_doc=doc, preset=preset)

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_area_pattern_from_preset", EventArgs.from_args(cargs))  # type: ignore
        return fe
