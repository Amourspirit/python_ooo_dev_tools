from __future__ import annotations
from typing import Tuple
import uno

from ooodev.format.inner.direct.chart2.series.data_series.borders.line_properties import (
    _LinePropertiesProps,
    LineProperties as SeriesLineProperties,
)
from ooodev.format.inner.preset.preset_border_line import BorderLineKind, get_preset_series_border_line_props


class LineProperties(SeriesLineProperties):
    """This class represents the line properties of a chart grid line properties."""

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
