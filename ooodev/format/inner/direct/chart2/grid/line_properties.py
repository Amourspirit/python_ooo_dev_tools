from __future__ import annotations
from typing import Tuple
import uno

from ooodev.format.inner.direct.chart2.series.data_series.borders.line_properties import (
    _LinePropertiesProps,
    LineProperties as SeriesLineProperties,
)
from ooodev.format.inner.preset.preset_border_line import BorderLineKind, get_preset_series_border_line_props
from ooodev.utils.color import Color
from ooodev.utils.data_type.intensity import Intensity
from ooodev.units.unit_obj import UnitT


class LineProperties(SeriesLineProperties):
    """
    This class represents the line properties of a chart grid line properties.

    .. seealso::

        - :ref:`help_chart2_format_direct_grid_line_properties`
    """

    def __init__(
        self,
        style: BorderLineKind = BorderLineKind.CONTINUOUS,
        color: Color = Color(0),
        width: float | UnitT = 0,
        transparency: int | Intensity = 0,
    ) -> None:
        """
        Constructor.

        Args:
            style (BorderLineKind): Line style. Defaults to ``BorderLineKind.CONTINUOUS``.
            color (Color, optional): Line Color. Defaults to ``Color(0)``.
            width (float | UnitT, optional): Line Width (in ``mm`` units) or :ref:`proto_unit_obj`. Defaults to ``0``.
            transparency (int | Intensity, optional): Line transparency from ``0`` to ``100``. Defaults to ``0``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_grid_line_properties`
        """
        super().__init__(style=style, color=color, width=width, transparency=transparency)

    def _set_line_style(self, style: BorderLineKind):
        props = get_preset_series_border_line_props(kind=style)
        self._set("LineCap", props.line_cap)
        self._set("LineDash", props.line_dash)
        self._set("LineDashName", props.line_dash_name)
        self._set("LineStyle", props.line_style)

    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.GridProperties",)
        return self._supported_services_values

    @property
    def _props(self) -> _LinePropertiesProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = _LinePropertiesProps(
                color1="LineColor",
                color2="",
                width="LineWidth",
                transparency1="LineTransparence",
                transparency2="",
            )
        return self._props_internal_attributes
