from __future__ import annotations
from typing import Any, Tuple
import uno

from ooo.dyn.chart.chart_axis_label_position import ChartAxisLabelPosition as ChartAxisLabelPosition

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase


class LabelPosition(StyleBase):
    """
    Chart Axis Label placement.

    .. seealso::

        - :ref:`help_chart2_format_direct_axis_positioning`

    .. versionadded:: 0.9.4
    """

    def __init__(self, pos: ChartAxisLabelPosition = ChartAxisLabelPosition.NEAR_AXIS) -> None:
        """
        Constructor

        Args:
            pos (ChartAxisLabelPosition, optional): Specifies where to place the labels: near axis, near axis (other side), outside start, or outside end.
                Default is ``ChartAxisLabelPosition.NEAR_AXIS``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_axis_positioning`
        """
        super().__init__()
        self.prop_pos = pos

    # region overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.Axis",)
        return self._supported_services_values

    # endregion overrides

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
    def prop_pos(self) -> ChartAxisLabelPosition:
        return self._get("LabelPosition")

    @prop_pos.setter
    def prop_pos(self, value: ChartAxisLabelPosition) -> None:
        self._set("LabelPosition", value)

    # endregion Properties
