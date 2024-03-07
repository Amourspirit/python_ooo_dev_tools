from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from ooo.dyn.awt.gradient_style import GradientStyle

from ooodev.format.inner.style_factory import area_transparency_gradient_factory
from ooodev.utils.data_type.intensity_range import IntensityRange
from ooodev.utils.data_type.offset import Offset
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.events.partial.events_partial import EventsPartial

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
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_area_transparency_gradient",
            after_event="after_style_area_transparency_gradient",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

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
        # pylint: disable=assignment-from-none
        doc = self._GradientPartial_transparency_get_chart_doc()
        styler = self.__styler
        factory = area_transparency_gradient_factory
        kwargs = {
            "style": style,
            "offset": offset,
            "angle": angle,
            "border": border,
            "grad_intensity": grad_intensity,
        }
        if doc is None:
            return styler.style(factory=factory, **kwargs)
        kwargs["chart_doc"] = doc
        return styler.style(factory=factory, **kwargs)

    def style_area_transparency_gradient_get(self) -> GradientT | None:
        """
        Gets the Area Transparency Gradient Style.

        Raises:
            CancelEventError: If the event ``before_style_area_transparency_gradient_get`` is cancelled and not handled.

        Returns:
            GradientT | None: Area transparency style or ``None`` if cancelled.
        """
        # pylint: disable=assignment-from-none
        doc = self._GradientPartial_transparency_get_chart_doc()
        styler = self.__styler
        return (
            styler.style_get(factory=area_transparency_gradient_factory)
            if doc is None
            else styler.style_get(factory=area_transparency_gradient_factory, chart_doc=doc)
        )
