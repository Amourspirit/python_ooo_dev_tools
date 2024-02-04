from __future__ import annotations
from typing import Any, Tuple, overload, NamedTuple
import uno
from com.sun.star.chart2 import XChartDocument

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo


class _AlignSeriesProps(NamedTuple):
    primary_y_axis: str
    has_y_axis: str
    has_secondary_y_axis: str


class AlignSeries(StyleBase):
    """
    Chart Data Series Align

    .. seealso::

        - :ref:`help_chart2_format_direct_series_series_options`
    """

    def __init__(
        self,
        chart_doc: XChartDocument,
        primary_y_axis: bool = True,
    ) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            primary_y_axis (bool, optional): If ``True`` Data Series is plotted on the primary Y axis; if ``False`` Data Series is plotted on the secondary Y axis. Defaults to ``True``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_series_series_options`
        """
        self._chart_doc = chart_doc
        super().__init__()
        if primary_y_axis is not None:
            self.prop_primary_y_axis = primary_y_axis

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.DataSeries",)
        return self._supported_services_values

    def _is_valid_obj(self, obj: Any) -> bool:
        return self._chart_doc is not None

    # region apply()
    @overload
    def apply(self, obj: Any) -> None: ...

    @overload
    def apply(self, obj: Any, **kwargs) -> None: ...

    def apply(self, obj: Any, **kwargs) -> None:
        """
        Applies styles to object

        Args:
            obj (object): UNO Object that styles are to be applied.

        Returns:
            None:
        """
        if not self._is_valid_obj(obj):
            self._print_not_valid_srv("apply")
            return

        dv = self._get_properties().copy()
        axis_idx = dv.pop(self._props.primary_y_axis, None)
        if axis_idx is not None:
            super().apply(obj=obj, validate=False, override_dv={self._props.primary_y_axis: axis_idx}, **kwargs)

        if dv:
            try:
                diagram = self._chart_doc.getDiagram()  # type: ignore
            except Exception as e:
                mLo.Lo.print(f"{self.__class__.__name__}.apply() - Unable to get chart diagram")
                mLo.Lo.print(f"  Error: {e}")
                return
            super().apply(obj=diagram, validate=False, override_dv=dv, **kwargs)

    # endregion apply()

    # region Copy()
    @overload
    def copy(self) -> AlignSeries: ...

    @overload
    def copy(self, **kwargs) -> AlignSeries: ...

    def copy(self, **kwargs) -> AlignSeries:
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
    def prop_primary_y_axis(self) -> bool | None:
        """Gets or sets the primary Y axis."""
        pv = self._get(self._props.primary_y_axis)
        return None if pv is None else pv == 0

    @prop_primary_y_axis.setter
    def prop_primary_y_axis(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.primary_y_axis)
            self._remove(self._props.has_y_axis)
            self._remove(self._props.has_secondary_y_axis)
            return
        self._set(self._props.primary_y_axis, 0 if value else 1)
        self._set(self._props.has_y_axis, value)
        self._set(self._props.has_secondary_y_axis, not value)

    @property
    def _props(self) -> _AlignSeriesProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = _AlignSeriesProps(
                primary_y_axis="AttachedAxisIndex",
                has_y_axis="HasYAxis",
                has_secondary_y_axis="HasSecondaryYAxis",
            )
        return self._props_internal_attributes

    # endregion Properties
