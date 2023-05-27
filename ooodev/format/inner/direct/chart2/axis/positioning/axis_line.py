from __future__ import annotations
from typing import Any, Tuple
import uno
from ooo.dyn.chart.chart_axis_position import ChartAxisPosition

from ooodev.events.args.key_val_cancel_args import KeyValCancelArgs
from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.utils.gen_util import NULL_OBJ, TNullObj


class AxisLine(StyleBase):
    """
    Chart Axis Line.

    Select where to cross the other axis: at start, at end, at a specified value, or at a category.

    .. seealso::

        - :ref:`help_chart2_format_direct_axis_positioning`

    .. versionadded:: 0.9.4
    """

    def __init__(self, cross: ChartAxisPosition | None = None, value: float | None | TNullObj = NULL_OBJ) -> None:  # type: ignore
        """
        Constructor

        Args:
            cross(ChartAxisPosition, optional): The position where the axis crosses the other axis.
            value (float, None, optional): The value where the axis crosses the other axis.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_axis_positioning`
        """
        # Because value can be None we use NULL_OBJ to indicate no value.
        super().__init__()
        if cross is not None:
            self.prop_cross = cross
        if value is not NULL_OBJ:
            self.prop_value = value

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.Axis",)
        return self._supported_services_values

    # endregion overrides

    # region On Events
    def on_property_setting(self, source: Any, event_args: KeyValCancelArgs) -> None:
        if event_args.key == "CrossoverValue" and event_args.value is None:
            # CrossoverValue can be None but is not allowed via property set.
            # Setting to default will cause Props to call set default.
            event_args.default = True
        return super().on_property_setting(source, event_args)

    # endregion On Events

    # region Properties
    @property
    def prop_format_kind(self) -> FormatKind:
        """Gets the kind of style"""
        try:
            return self._format_kind_prop
        except AttributeError:
            self._format_kind_prop = FormatKind.UNKNOWN
        return self._format_kind_prop

    @property
    def prop_cross(self) -> ChartAxisPosition | None:
        return self._get("CrossoverPosition")

    @prop_cross.setter
    def prop_cross(self, value: ChartAxisPosition | None) -> None:
        if value is None:
            self._remove("CrossoverPosition")
            return
        self._set("CrossoverPosition", value)

    @property
    def prop_value(self) -> Any:
        return self._get("CrossoverValue")

    @prop_value.setter
    def prop_value(self, value: Any) -> None:
        if value is NULL_OBJ:
            self._remove("CrossoverValue")
            return
        if value is None:
            self._set("CrossoverValue", None)
        else:
            self._set("CrossoverValue", float(value))

    # endregion Properties
