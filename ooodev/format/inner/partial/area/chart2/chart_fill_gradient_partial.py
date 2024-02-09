from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooo.dyn.awt.gradient_style import GradientStyle
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_factory import chart2_area_gradient_factory
from ooodev.loader import lo as mLo
from ooodev.utils import color as mColor
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.data_type.color_range import ColorRange
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.data_type.offset import Offset

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.area.chart2.chart_fill_gradient_t import ChartFillGradientT
    from ooodev.units import Angle
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.format.inner.preset.preset_gradient import PresetGradientKind
else:
    XChartDocument = Any
    Angle = Any
    Intensity = Any
    PresetGradientKind = Any


class ChartFillGradientPartial:
    """
    Partial class for Chart Fill Gradient.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

    def _ChartFillGradientPartial__get_chart_doc(self) -> XChartDocument:
        raise NotImplementedError

    def style_area_gradient(
        self,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        step_count: int = 0,
        offset: Offset = Offset(50, 50),
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        grad_color: ColorRange = ColorRange(mColor.Color(0), mColor.Color(16777215)),
        grad_intensity: IntensityRange = IntensityRange(100, 100),
        name: str = "",
    ) -> ChartFillGradientT | None:
        """
        Style Area Color.

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (Offset, int, optional): Specifies the X and Y coordinate, where the gradient begins.
                X is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and ``RECT``
                style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to ``0``.
            border (int, optional): Specifies percent of the total width where just the start color is used.
                Defaults to ``0``.
            grad_color (ColorRange, optional): Specifies the color at the start point and stop point of the gradient.
                Defaults to ``ColorRange(Color(0), Color(16777215))``.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of the
                gradient. Defaults to ``IntensityRange(100, 100)``.
            name (str, optional): Specifies the Fill Gradient Name.

        Raises:
            CancelEventError: If the event ``before_style_area_gradient`` is cancelled and not handled.

        Returns:
            ChartFillGradientT | None: Chart Fill Gradient instance or ``None`` if cancelled.
        """
        doc = self._ChartFillGradientPartial__get_chart_doc()
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_area_gradient.__qualname__)
            event_data: Dict[str, Any] = {
                "style": style,
                "step_count": step_count,
                "offset": offset,
                "angle": angle,
                "border": border,
                "grad_color": grad_color,
                "grad_intensity": grad_intensity,
                "name": name,
                "factory_name": factory_name,
                "component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_area_gradient", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_area_gradient")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style Area Gradient has been cancelled.")
                    else:
                        return None
                else:
                    return None
            style = cargs.event_data.get("style", style)
            step_count = cargs.event_data.get("step_count", step_count)
            offset = cargs.event_data.get("offset", offset)
            angle = cargs.event_data.get("angle", angle)
            border = cargs.event_data.get("border", border)
            grad_color = cargs.event_data.get("grad_color", grad_color)
            grad_intensity = cargs.event_data.get("grad_intensity", grad_intensity)
            name = cargs.event_data.get("name", name)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("component", comp)

        styler = chart2_area_gradient_factory(factory_name)
        fe = styler(
            chart_doc=doc,
            style=style,
            step_count=step_count,
            offset=offset,
            angle=angle,
            border=border,
            grad_color=grad_color,
            grad_intensity=grad_intensity,
            name=name,
        )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        if has_events:
            self.trigger_event("after_style_area_gradient", EventArgs.from_args(cargs))  # type: ignore
        return fe

    def style_area_gradient_from_preset(self, preset: PresetGradientKind) -> ChartFillGradientT | None:
        """
        Style Area Gradient from Preset.

        Args:
            preset (PresetGradientKind): Preset Gradient Kind.

        Returns:
            ChartFillGradientT: Chart Fill Gradient instance.
        """
        doc = self._ChartFillGradientPartial__get_chart_doc()
        cargs = None
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        comp = self.__component
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_area_gradient.__qualname__)
            event_data: Dict[str, Any] = {
                "preset": preset,
                "factory_name": factory_name,
                "component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_area_gradient_from_preset", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_area_gradient_from_preset")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style Area Gradient has been cancelled.")
                    else:
                        return None
                else:
                    return None
            preset = cargs.event_data.get("preset", preset)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("component", comp)

        styler = chart2_area_gradient_factory(factory_name)
        fe = styler.from_preset(chart_doc=doc, preset=preset)

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        if has_events:
            self.trigger_event("after_style_area_gradient_from_preset", EventArgs.from_args(cargs))  # type: ignore
        return fe
