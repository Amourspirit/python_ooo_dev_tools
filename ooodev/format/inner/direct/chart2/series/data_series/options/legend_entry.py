from __future__ import annotations
from typing import Tuple, overload
import uno
from com.sun.star.chart2 import XChartDocument

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase


class LegendEntry(StyleBase):
    """
    Chart Data Series Legend Visibility

    .. seealso::

        - :ref:`help_chart2_format_direct_series_series_options`
    """

    def __init__(self, chart_doc: XChartDocument, *, hide_legend: bool = False, **kwargs) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            hide_legend (bool, optional): Specifies if legend entry is to be hidden. Defaults to ``False``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_series_series_options`
        """
        self._chart_doc = chart_doc
        super().__init__()
        self.prop_hide_legend = hide_legend

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = (
                "com.sun.star.chart2.DataSeries",
                "com.sun.star.chart2.DataPointProperties",
            )
        return self._supported_services_values

    # region apply()
    @overload
    def apply(self, obj: object) -> None:
        ...

    @overload
    def apply(self, obj: object, **kwargs) -> None:
        ...

    def apply(self, obj: object, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO Object that styles are to be applied.

        Returns:
            None:
        """
        # obj is data series object.
        if not self._is_valid_obj(obj):
            self._print_not_valid_srv("apply")
            return
        super().apply(obj=obj, validate=False, override_dv={"ShowLegendEntry": not self.prop_hide_legend})

    # endregion apply()

    # region Copy()
    @overload
    def copy(self) -> LegendEntry:
        ...

    @overload
    def copy(self, **kwargs) -> LegendEntry:
        ...

    def copy(self, **kwargs) -> LegendEntry:
        """Gets a copy of instance as a new instance"""
        cp = super().copy(**kwargs)
        cp._chart_doc = self._chart_doc
        return cp

    # endregion Copy()

    # endregion Overrides

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
    def prop_hide_legend(self) -> bool:
        """Gets or sets whether the legend is hidden."""
        pv = self._get("ShowLegendEntry")
        return not pv

    @prop_hide_legend.setter
    def prop_hide_legend(self, value: bool) -> None:
        self._set("ShowLegendEntry", not value)

    # endregion Properties
