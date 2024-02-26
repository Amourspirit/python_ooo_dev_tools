from __future__ import annotations
import contextlib
from typing import Any
import uno
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs

from ooodev.format.inner.direct.chart2.series.data_series.borders.line_properties import (
    LineProperties as DataSeriesLineProperties,
    _LinePropertiesProps,
)
from ooodev.format.inner.preset.preset_border_line import BorderLineKind, get_preset_series_border_line_props
from ooodev.units.unit_obj import UnitT
from ooodev.utils import props as mProps
from ooodev.utils.color import Color
from ooodev.utils.data_type.intensity import Intensity


class LineProperties(DataSeriesLineProperties):
    """
    This class represents the line properties of a chart data series labels borders line properties.

    .. seealso::

        - :ref:`help_chart2_format_direct_series_labels_borders`
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
            width (float, UnitT, optional): Line Width (in ``mm`` units) or :ref:`proto_unit_obj`. Defaults to ``0``.
            transparency (int, Intensity, optional): Line transparency from ``0`` to ``100``. Defaults to ``0``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_series_labels_borders`
        """
        super().__init__(style=style, color=color, width=width, transparency=transparency)

    # region overrides
    def _set_line_style(self, style: BorderLineKind):
        props = get_preset_series_border_line_props(kind=style)

        label_border_dash = mProps.Props.make_props(Name=props.border_name, LineDash=props.line_dash)
        self._set("LabelBorderDashName", props.border_name)
        self._set("LabelBorderStyle", props.border_style)
        self._set("LabelBorderDash", label_border_dash)

    # endregion overrides

    # region event methods
    def on_property_set_error(self, source: Any, event_args: KeyValCancelArgs) -> None:
        if event_args.key == "LabelBorderDash":
            # there is a bug, API DataPoint does not properly implement XPropertySet.
            # This may lead to an exception when trying to setting the LabelBorderDash property.
            # Even though the property is set, the border appears on the chart data point label.
            # Opening the data point properties dialog shows the border line is not set to anything.
            # However, because it display correctly, we will ignore the exception.
            with contextlib.suppress(Exception):
                if event_args.event_data.getImplementationName() == "com.sun.star.comp.chart.DataPoint":
                    event_args.handled = True
                    # attempting to set via invoke does not work, setPropertyValue is missing due to bad implementation.
                    #
                    # tpl = self._get("LabelBorderDash")
                    # uno_any = uno.Any("[]com.sun.star.beans.XPropertySet", tpl)
                    # props = mLo.Lo.qi(XPropertySet, event_args.event_data)
                    # uno.invoke(props, "setPropertyValue", ("LabelBorderDash", uno_any))
                    # event_args.handled = True
        return super().on_property_set_error(source, event_args)

    # endregion event methods

    # region Properties
    @property
    def _props(self) -> _LinePropertiesProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = _LinePropertiesProps(
                color1="LabelBorderColor",
                color2="",
                width="LabelBorderWidth",
                transparency1="LabelBorderTransparency",
                transparency2="",
            )
        return self._props_internal_attributes

    # endregion Properties
