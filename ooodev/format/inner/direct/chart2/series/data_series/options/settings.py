from __future__ import annotations
from typing import Any, Tuple, cast, overload, NamedTuple
import uno
from com.sun.star.chart2 import XChartDocument
from com.sun.star.beans import XPropertySet

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.loader import lo as mLo
from ooodev.utils import props as mProps


class _SettingsProps(NamedTuple):
    primary_y_axis: str
    spacing: str
    overlap: str
    side_by_side: str


class Settings(StyleBase):
    """
    Chart Data Series Settings

    Note:
        The axis that the setting are applied to is determined by the axis that the data series is plotted on.
        For this reason if formatting is applied to a data series axis it should be done before applying ``Settings``.

    .. seealso::

        - :ref:`help_chart2_format_direct_series_series_options`

    .. versionadded:: 0.9.4
    """

    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        spacing: int | None = None,
        overlap: int | None = None,
        side_by_side: bool | None = None,
        **kwargs,
    ) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            spacing (int | None, optional): Spacing between bars. Must be a positive integer value. Defaults to ``None``.
            overlap (int | None, optional): Overlap of bars. Defaults to ``None``.
            side_by_side (bool | None, optional): Specifies if bars are shown side by side. Defaults to ``None``.

        Returns:
            None:

        See Also:
            - :ref:`help_chart2_format_direct_series_series_options`
        """
        self._chart_doc = chart_doc
        super().__init__()
        if spacing is not None:
            self.prop_spacing = spacing
        if overlap is not None:
            self.prop_overlap = overlap
        if side_by_side is not None:
            self.prop_side_by_side = side_by_side

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ("com.sun.star.chart2.DataSeries",)
        return self._supported_services_values

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
        # obj is a data series object.
        if not self._is_valid_obj(obj):
            self._print_not_valid_srv("apply")
            return

        try:
            diagram = self._chart_doc.getDiagram()  # type: ignore
        except Exception as e:
            mLo.Lo.print(f"{self.__class__.__name__}.apply() - Unable to apply spacing and overlap.")
            mLo.Lo.print(f"  Error: {e}")
            return

        if self.prop_side_by_side is not None:
            super().apply(
                obj=diagram, validate=False, override_dv={self._props.side_by_side: self.prop_side_by_side}, **kwargs
            )

        spacing = self.prop_spacing
        overlap = self.prop_overlap
        axis_props = {}
        axis = None
        if spacing is not None or overlap is not None:
            try:
                # get the axis to apply spacing and overlap.
                # The axis is determined by the primary_y_axis property of the data series.
                # if primary_y_axis is True, then the YAxis is used, otherwise the SecondaryYAxis is used.
                is_primary_y = int(mProps.Props.get(obj, self._props.primary_y_axis, 0)) == 0
                if is_primary_y:
                    axis = cast(XPropertySet, mProps.Props.get(diagram, "YAxis"))
                else:
                    axis = cast(XPropertySet, mProps.Props.get(diagram, "SecondaryYAxis"))

                if spacing is not None:
                    axis_props[self._props.spacing] = spacing
                if overlap is not None:
                    axis_props[self._props.overlap] = overlap
            except Exception as e:
                mLo.Lo.print(f"{self.__class__.__name__}.apply() - Unable to apply spacing and overlap.")
                mLo.Lo.print(f"  Error: {e}")
                return
        if axis_props and axis is not None:
            super().apply(obj=axis, validate=False, override_dv=axis_props, **kwargs)

    # endregion apply()
    # region Copy()
    @overload
    def copy(self) -> Settings: ...

    @overload
    def copy(self, **kwargs) -> Settings: ...

    def copy(self, **kwargs) -> Settings:
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
    def prop_spacing(self) -> int | None:
        """Gets or sets the spacing between bars."""
        return self._get(self._props.spacing)

    @prop_spacing.setter
    def prop_spacing(self, value: int | None) -> None:
        if value is None:
            self._remove(self._props.spacing)
            return
        value = max(value, 0)
        self._set(self._props.spacing, value)

    @property
    def prop_overlap(self) -> int | None:
        """Gets or sets the overlap between bars."""
        return self._get(self._props.overlap)

    @prop_overlap.setter
    def prop_overlap(self, value: int | None) -> None:
        if value is None:
            self._remove(self._props.overlap)
            return
        self._set(self._props.overlap, value)

    @property
    def prop_side_by_side(self) -> bool | None:
        """Gets or sets whether bars are side by side."""
        pv = cast(bool, self._get(self._props.side_by_side))
        # GroupBarsPerAxis is the opposite of SideBySide
        return None if pv is None else not pv

    @prop_side_by_side.setter
    def prop_side_by_side(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.side_by_side)
            return
        # GroupBarsPerAxis is the opposite of SideBySide
        self._set(self._props.side_by_side, not value)

    @property
    def _props(self) -> _SettingsProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = _SettingsProps(
                primary_y_axis="AttachedAxisIndex",
                spacing="GapWidth",
                overlap="Overlap",
                side_by_side="GroupBarsPerAxis",
            )
        return self._props_internal_attributes

    # endregion Properties
