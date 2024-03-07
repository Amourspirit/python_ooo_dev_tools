from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooo.dyn.awt.gradient_style import GradientStyle
from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.format.inner.style_factory import chart2_area_gradient_factory
from ooodev.utils import color as mColor
from ooodev.utils.data_type.color_range import ColorRange
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.data_type.offset import Offset
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.chart2.area.chart_fill_gradient_t import ChartFillGradientT
    from ooodev.units.angle import Angle
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
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_area_gradient",
            after_event="after_style_area_gradient",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def _ChartFillGradientPartial__get_chart_doc(self) -> XChartDocument:
        if isinstance(self, ChartDocPropPartial):
            return self.chart_doc.component
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

        Hint:
            - ``GradientStyle`` can be imported from ``ooo.dyn.awt.gradient_style``
            - ``Angle`` can be imported from ``ooodev.units``
            - ``ColorRange`` can be imported from ``ooodev.utils.data_type.color_range``
            - ``Intensity`` can be imported from ``ooodev.utils.data_type.intensity``
            - ``IntensityRange`` can be imported from ``ooodev.utils.data_type.intensity_range``
            - ``Offset`` can be imported from ``ooodev.utils.data_type.offset``
        """
        doc = self._ChartFillGradientPartial__get_chart_doc()
        factory = chart2_area_gradient_factory
        kwargs = {
            "chart_doc": doc,
            "style": style,
            "step_count": step_count,
            "offset": offset,
            "angle": angle,
            "border": border,
            "grad_color": grad_color,
            "grad_intensity": grad_intensity,
            "name": name,
        }
        return self.__styler.style(factory=factory, **kwargs)

    def style_area_gradient_get(self) -> ChartFillGradientT | None:
        """
        Gets the Area Gradient Style.

        Raises:
            CancelEventError: If the event ``before_style_area_gradient_get`` is cancelled and not handled.

        Returns:
            ChartFillGradientT | None: Gradient style or ``None`` if cancelled.
        """
        doc = self._ChartFillGradientPartial__get_chart_doc()
        return self.__styler.style_get(factory=chart2_area_gradient_factory, chart_doc=doc)

    def style_area_gradient_from_preset(self, preset: PresetGradientKind) -> ChartFillGradientT | None:
        """
        Style Area Gradient from Preset.

        Args:
            preset (PresetGradientKind): Preset Gradient Kind.

        Returns:
            ChartFillGradientT: Chart Fill Gradient instance.

        Hint:
            - ``PresetGradientKind`` can be imported from ``ooodev.format.inner.preset.preset_gradient``
        """
        styler = self.__styler
        doc = self._ChartFillGradientPartial__get_chart_doc()
        fe = styler.style_get(
            factory=chart2_area_gradient_factory,
            call_method_name="from_preset",
            event_name_suffix="_from_preset",
            obj_arg_name="",
            chart_doc=doc,
            preset=preset,
        )
        styler.style_apply(style=fe, chart_doc=doc, preset=preset)
        return fe
