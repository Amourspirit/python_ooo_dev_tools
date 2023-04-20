from __future__ import annotations
from typing import Any
import uno
from com.sun.star.beans import XPropertySet
from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs

from ooodev.format.inner.direct.chart2.series.data_series.borders.line_properties import (
    LineProperties as DataSeriesLineProperties,
    _LinePropertiesProps,
)
from ooodev.format.inner.preset.preset_border_line import BorderLineKind, get_preset_series_border_line_props
from ooodev.utils import props as mProps
from ooodev.utils import lo as mLo


class LineProperties(DataSeriesLineProperties):
    """This class represents the line properties of a chart data series labels borders line properties."""

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
            # This may lead to an exception when trying to setint the LabelBorderDash property.
            # Even though the property is set, the border appears on the chart data point lablel.
            # Opening the data point properties dialog shows the border line is not set to anything.
            # However, because it display correctly, we will ignore the exception.
            try:
                if event_args.event_data.getImplementationName() == "com.sun.star.comp.chart.DataPoint":
                    event_args.handled = True
                    # attempting to set via invoke does not work, setPropertyValue is missing due to bad implementation.
                    #
                    # tpl = self._get("LabelBorderDash")
                    # uany = uno.Any("[]com.sun.star.beans.XPropertySet", tpl)
                    # props = mLo.Lo.qi(XPropertySet, event_args.event_data)
                    # uno.invoke(props, "setPropertyValue", ("LabelBorderDash", uany))
                    # event_args.handled = True
            except Exception:
                pass
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
