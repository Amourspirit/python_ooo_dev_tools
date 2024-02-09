from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING
import uno
from ooo.dyn.drawing.hatch_style import HatchStyle

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_factory import chart2_area_hatch_factory
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils import color as mColor

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
    from ooodev.format.inner.preset.preset_hatch import PresetHatchKind
    from ooodev.format.proto.area.chart2.chart_fill_hatch_t import ChartFillHatchT
    from ooodev.units import UnitT
    from ooodev.units import Angle
    from ooodev.loader.inst.lo_inst import LoInst
else:
    XChartDocument = Any
    PresetHatchKind = Any
    ChartFillHatchT = Any
    UnitT = Any
    Angle = Any
    LoInst = Any


class ChartFillHatchPartial:
    """
    Partial class for Chart Fill Pattern.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

    def _ChartFillHatchPartial__get_chart_doc(self) -> XChartDocument:
        raise NotImplementedError

    def style_area_hatch(
        self,
        *,
        style: HatchStyle = HatchStyle.SINGLE,
        color: mColor.Color = mColor.Color(0),
        space: float | UnitT = 0,
        angle: Angle | int = 0,
        bg_color: mColor.Color = mColor.Color(-1),
    ) -> ChartFillHatchT | None:
        """
        Style Area Color.

        Args:
            style (HatchStyle, optional): Specifies the kind of lines used to draw this hatch. Default ``HatchStyle.SINGLE``.
            color (:py:data:`~.utils.color.Color`, optional): Specifies the color of the hatch lines. Default ``0``.
            space (float, UnitT, optional): Specifies the space between the lines in the hatch (in ``mm`` units) or :ref:`proto_unit_obj`. Default ``0.0``
            angle (Angle, int, optional): Specifies angle of the hatch in degrees. Default to ``0``.
            bg_color(Color, optional): Specifies the background Color. Set this ``-1`` (default) for no background color.

        Raises:
            CancelEventError: If the event ``before_style_area_hatch`` is cancelled and not handled.

        Returns:
            ChartFillHatchT | None: Fill Image instance or ``None`` if cancelled.
        """
        doc = self._ChartFillHatchPartial__get_chart_doc()
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_area_hatch.__qualname__)
            event_data: Dict[str, Any] = {
                "style": style,
                "color": color,
                "space": space,
                "angle": angle,
                "bg_color": bg_color,
                "factory_name": factory_name,
                "component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_area_hatch", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_area_hatch")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style Area Hatch has been cancelled.")
                    else:
                        return None
                else:
                    return None
            style = cargs.event_data.get("style", style)
            color = cargs.event_data.get("color", color)
            space = cargs.event_data.get("space", space)
            angle = cargs.event_data.get("angle", angle)
            bg_color = cargs.event_data.get("bg_color", bg_color)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("component", comp)

        styler = chart2_area_hatch_factory(factory_name)
        fe = styler(
            chart_doc=doc,
            style=style,
            color=color,
            space=space,
            angle=angle,
            bg_color=bg_color,
        )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_area_hatch", EventArgs.from_args(cargs))  # type: ignore
        return fe

    def style_area_hatch_from_preset(self, preset: PresetHatchKind) -> ChartFillHatchT | None:
        """
        Style Area Gradient from Preset.

        Args:
            preset (PresetHatchKind): Preset Hatch Kind.

        Returns:
            ChartFillHatchT | None: Chart Fill Hatch instance or ``None`` if ``before_style_area_hatch_from_preset`` event is cancelled.
        """
        doc = self._ChartFillHatchPartial__get_chart_doc()
        cargs = None
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        comp = self.__component
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_area_hatch_from_preset.__qualname__)
            event_data: Dict[str, Any] = {
                "preset": preset,
                "factory_name": factory_name,
                "component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_area_hatch_from_preset", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_area_hatch_from_preset")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style Area Hatch has been cancelled.")
                    else:
                        return None
                else:
                    return None
            preset = cargs.event_data.get("preset", preset)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("component", comp)

        styler = chart2_area_hatch_factory(factory_name)
        fe = styler.from_preset(chart_doc=doc, preset=preset)

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_area_hatch_from_preset", EventArgs.from_args(cargs))  # type: ignore
        return fe
