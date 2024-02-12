from __future__ import annotations
from typing import Any, Dict, TYPE_CHECKING

from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_factory import area_transparency_gradient_factory
from ooodev.loader import lo as mLo
from ooodev.utils.context.lo_context import LoContext

from ooo.dyn.awt.gradient_style import GradientStyle
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.data_type.offset import Offset

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.units import Angle
    from ooodev.utils.data_type.intensity import Intensity
    from ooodev.format.proto.area.transparency.gradient_t import GradientT
else:
    XChartDocument = Any
    GradientT = Any
    Angle = Any
    Intensity = Any


class GradientPartial:
    """
    Partial class for FillColor.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__lo_inst = lo_inst
        self.__factory_name = factory_name
        self.__component = component

    def _GradientPartial_transparency_get_chart_doc(self) -> XChartDocument | None:
        return None

    def style_area_transparency_gradient(
        self,
        *,
        style: GradientStyle = GradientStyle.LINEAR,
        offset: Offset = Offset(50, 50),
        angle: Angle | int = 0,
        border: Intensity | int = 0,
        grad_intensity: IntensityRange = IntensityRange(0, 0),
    ) -> GradientT | None:
        """
        Style Area Color.

        Args:
            style (GradientStyle, optional): Specifies the style of the gradient. Defaults to ``GradientStyle.LINEAR``.
            step_count (int, optional): Specifies the number of steps of change color. Defaults to ``0``.
            offset (offset, optional): Specifies the X-coordinate (start) and Y-coordinate (end),
                where the gradient begins. X is effectively the center of the ``RADIAL``, ``ELLIPTICAL``, ``SQUARE`` and
                ``RECT`` style gradients. Defaults to ``Offset(50, 50)``.
            angle (Angle, int, optional): Specifies angle of the gradient. Defaults to ``0``.
            border (int, optional): Specifies percent of the total width where just the start color is used.
                Defaults to ``0``.
            grad_intensity (IntensityRange, optional): Specifies the intensity at the start point and stop point of
                the gradient. Defaults to ``IntensityRange(0, 0)``.

        Raises:
            CancelEventError: If the event ``before_style_area_transparency_gradient`` is cancelled and not handled.

        Returns:
            GradientT | None: FillColor instance or ``None`` if cancelled.

        Hint:
            - ``GradientStyle`` can be imported from ``ooo.dyn.awt.gradient_style``
            - ``IntensityRange`` can be imported from ``ooodev.utils.data_type.intensity_range``
            - ``Offset`` can be imported from ``ooodev.utils.data_type.offset``
            - ``Angle`` can be imported from ``ooodev.units``
            - ``Intensity`` can be imported from ``ooodev.utils.data_type.intensity``
        """
        doc = self._GradientPartial_transparency_get_chart_doc()
        comp = self.__component
        factory_name = self.__factory_name
        has_events = False
        cargs = None
        if isinstance(self, EventsPartial):
            has_events = True
            cargs = CancelEventArgs(self.style_area_transparency_gradient.__qualname__)
            event_data: Dict[str, Any] = {
                "style": style,
                "offset": offset,
                "angle": angle,
                "border": border,
                "grad_intensity": grad_intensity,
                "factory_name": factory_name,
                "this_component": comp,
            }
            cargs.event_data = event_data
            self.trigger_event("before_style_area_transparency_gradient", cargs)
            if cargs.cancel is True:
                if cargs.handled is False:
                    cargs.set("initial_event", "before_style_area_transparency_gradient")
                    self.trigger_event(GblNamedEvent.EVENT_CANCELED, cargs)
                    if cargs.handled is False:
                        raise mEx.CancelEventError(cargs, "Style Font Effects has been cancelled.")
                    else:
                        return None
                else:
                    return None
            style = cargs.event_data.get("style", style)
            offset = cargs.event_data.get("offset", offset)
            angle = cargs.event_data.get("angle", angle)
            border = cargs.event_data.get("border", border)
            grad_intensity = cargs.event_data.get("grad_intensity", grad_intensity)
            factory_name = cargs.event_data.get("factory_name", factory_name)
            comp = cargs.event_data.get("this_component", comp)

        styler = area_transparency_gradient_factory(factory_name)
        if doc is None:
            fe = styler(style=style, offset=offset, angle=angle, border=border, grad_intensity=grad_intensity)
        else:
            fe = styler(
                chart_doc=doc, style=style, offset=offset, angle=angle, border=border, grad_intensity=grad_intensity
            )

        if has_events:
            fe.add_event_observer(self.event_observer)  # type: ignore

        with LoContext(self.__lo_inst):
            fe.apply(comp)
        fe.set_update_obj(comp)
        if has_events:
            self.trigger_event("after_style_area_transparency_gradient", EventArgs.from_args(cargs))  # type: ignore
        return fe
