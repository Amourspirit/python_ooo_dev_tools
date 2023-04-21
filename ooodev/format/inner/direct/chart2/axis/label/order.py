from __future__ import annotations
from typing import cast, Tuple
import uno
from ooo.dyn.chart.chart_axis_arrange_order_type import ChartAxisArrangeOrderType

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase


class Order(StyleBase):
    """
    Chart Axis Label visibility.

    .. versionadded:: 0.9.4
    """

    def __init__(self, order: ChartAxisArrangeOrderType = ChartAxisArrangeOrderType.AUTO) -> None:
        """
        Constructor

        Args:
            order (ChartAxisArrangeOrderType, optional): Specifies the arrangement of the axes descriptions.

        Returns:
            None:
        """
        super().__init__()
        self.prop_order = order

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
    def prop_order(self) -> ChartAxisArrangeOrderType:
        """Gets or Sets if the Axis Label is visible."""
        return cast(ChartAxisArrangeOrderType, self._get("ArrangeOrder"))

    @prop_order.setter
    def prop_order(self, value: ChartAxisArrangeOrderType) -> None:
        self._set("ArrangeOrder", value)

    # endregion Properties
