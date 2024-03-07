from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from ooo.dyn.drawing.hatch_style import HatchStyle

from ooodev.calc.chart2.partial.chart_doc_prop_partial import ChartDocPropPartial
from ooodev.format.inner.style_factory import chart2_area_hatch_factory
from ooodev.format.inner.partial.default_factor_styler import DefaultFactoryStyler
from ooodev.utils import color as mColor
from ooodev.events.partial.events_partial import EventsPartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import XChartDocument
    from ooodev.format.inner.preset.preset_hatch import PresetHatchKind
    from ooodev.format.proto.chart2.area.chart_fill_hatch_t import ChartFillHatchT
    from ooodev.units.unit_obj import UnitT
    from ooodev.units.angle import Angle
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
        self.__styler = DefaultFactoryStyler(
            factory_name=factory_name,
            component=component,
            before_event="before_style_area_hatch",
            after_event="after_style_area_hatch",
            lo_inst=lo_inst,
        )
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)

    def _ChartFillHatchPartial__get_chart_doc(self) -> XChartDocument:
        if isinstance(self, ChartDocPropPartial):
            return self.chart_doc.component

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

        Hint:
            - ``Angle`` can be imported from ``ooodev.units``
            - ``HatchStyle`` can be imported from ``ooo.dyn.drawing.hatch_style``
        """
        doc = self._ChartFillHatchPartial__get_chart_doc()
        factory = chart2_area_hatch_factory
        kwargs = {
            "chart_doc": doc,
            "style": style,
            "color": color,
            "space": space,
            "angle": angle,
            "bg_color": bg_color,
        }
        return self.__styler.style(factory=factory, **kwargs)

    def style_area_hatch_from_preset(self, preset: PresetHatchKind) -> ChartFillHatchT | None:
        """
        Style Area Gradient from Preset.

        Args:
            preset (PresetHatchKind): Preset Hatch Kind.

        Returns:
            ChartFillHatchT | None: Chart Fill Hatch instance or ``None`` if ``before_style_area_hatch_from_preset`` event is cancelled.

        Hint:
            - ``PresetHatchKind`` can be imported from ``ooodev.format.inner.preset.preset_hatch``
        """
        styler = self.__styler
        doc = self._ChartFillHatchPartial__get_chart_doc()
        fe = styler.style_get(
            factory=chart2_area_hatch_factory,
            call_method_name="from_preset",
            event_name_suffix="_from_preset",
            obj_arg_name="",
            chart_doc=doc,
            preset=preset,
        )
        styler.style_apply(style=fe, chart_doc=doc, preset=preset)
        return fe
