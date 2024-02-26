from __future__ import annotations
from typing import Any, TYPE_CHECKING

from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.preset.preset_border_line import BorderLineKind
from ooodev.loader import lo as mLo
from ooodev.utils import color as mColor
from ooodev.format.inner.partial.draw.borders.line_properties import LineProperties

if TYPE_CHECKING:
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.format.proto.borders.line_properties_t import LinePropertiesT
    from ooodev.units.unit_obj import UnitT
    from ooodev.utils.data_type.intensity import Intensity
else:
    LoInst = Any
    LinePropertiesT = Any
    UnitT = Any
    Intensity = Any


class AxisLinePropertiesPartial:
    """
    Partial class for Chart2 Axis Line Properties.
    """

    def __init__(self, factory_name: str, component: Any, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self.__styler = LineProperties(factory_name=factory_name, component=component, lo_inst=lo_inst)
        if isinstance(self, EventsPartial):
            self.__styler.add_event_observers(self.event_observer)
        self.__styler.after_event_name = "after_style_axis_line"
        self.__styler.before_event_name = "before_style_axis_line"

    def style_axis_line(
        self,
        color: mColor.Color = mColor.Color(0),
        width: float | UnitT = 0,
        transparency: int | Intensity = 0,
        style: BorderLineKind = BorderLineKind.CONTINUOUS,
    ) -> LinePropertiesT | None:
        """
        Style Axis Line.

        Args:
            color (Color, optional): Line Color. Defaults to ``Color(0)``.
            width (float | UnitT, optional): Line Width (in ``mm`` units) or :ref:`proto_unit_obj`. Defaults to ``0``.
            transparency (int | Intensity, optional): Line transparency from ``0`` to ``100``. Defaults to ``0``.
            style (BorderLineKind, optional): Line style. Defaults to ``BorderLineKind.CONTINUOUS``.

        Raises:
            CancelEventError: If the event ``before_style_axis_line`` is cancelled and not handled.

        Returns:
            LinePropertiesT | None: Axis Line Style or ``None`` if cancelled.

        Hint:
            - ``BorderLineKind`` can be imported from ``ooodev.format.inner.preset.preset_border_line``
            - ``Intensity`` can be imported from ``ooodev.utils.data_type.intensity``
        """
        return self.__styler.style(color=color, width=width, transparency=transparency, style=style)

    def style_axis_line_get(self) -> LinePropertiesT | None:
        """
        Gets the axis line properties style.

        Raises:
            CancelEventError: If the event ``before_style_axis_line_get`` is cancelled and not handled.

        Returns:
            LinePropertiesT | None: Line properties style or ``None`` if cancelled.
        """
        return self.__styler.style_get()
