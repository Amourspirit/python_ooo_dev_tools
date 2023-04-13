from __future__ import annotations
from typing import Tuple, cast, overload, NamedTuple, TYPE_CHECKING, Union
import uno
from com.sun.star.beans import XPropertySet
from com.sun.star.chart2 import XChartDocument
from com.sun.star.chart2 import XDiagramProvider

from ooo.dyn.chart.missing_value_treatment import MissingValueTreatmentEnum

from ooodev.format.inner.kind.format_kind import FormatKind
from ooodev.format.inner.style_base import StyleBase
from ooodev.office.chart2 import Chart2
from ooodev.utils import lo as mLo
from ooodev.utils import props as mProps

if TYPE_CHECKING:
    from com.sun.star.chart2 import Diagram  # service
    from com.sun.star.chart2.data import DataProvider  # service


class _OptionsProps(NamedTuple):
    primary_y_axis: str
    spacing: str
    overlap: str
    side_by_side: str
    missing_values: str
    hidden_cell_values: str
    show_legend: str
    has_y_axis: str
    has_secondary_y_axis: str


class Options(StyleBase):
    def __init__(
        self,
        chart_doc: XChartDocument,
        *,
        primary_y_axis: bool | None = None,
        spacing: int | None = None,
        overlap: int | None = None,
        side_by_side: bool | None = None,
        missing_values: MissingValueTreatmentEnum | None = None,
        hidden_cell_values: bool | None = None,
        hide_legend: bool | None = None,
    ) -> None:
        """
        Constructor

        Args:
            chart_doc (XChartDocument): Chart document.
            primary_y_axis (bool | None, optional): If ``True`` Data Series is plotted on the primary Y axis; if ``False`` Data Series is plotted on the secondary Y axis. Defaults to ``None``.
            spacing (int | None, optional): Spacing between bars. Must be a positive integer value. Defaults to ``None``.
            overlap (int | None, optional): Overlap of bars. Defaults to ``None``.
            side_by_side (bool | None, optional): Specifies if bars are shown side by side. Defaults to ``None``.
            missing_values (MissingValueTreatmentEnum | None, optional): Specifies plot missing values option. Defaults to ``None``.
            hidden_cell_values (bool | None, optional): Specifies if values from hidden cells are to be included. Defaults to ``None``.
            hide_legend (bool | None, optional): Specifies if legend entry is to be hidden. Defaults to ``None``.
        """
        self._chart_doc = chart_doc
        super().__init__()
        if primary_y_axis is not None:
            self.prop_primary_y_axis = primary_y_axis
        if spacing is not None:
            self.prop_spacing = spacing
        if overlap is not None:
            self.prop_overlap = overlap
        if side_by_side is not None:
            self.prop_side_by_side = side_by_side
        if missing_values is not None:
            self.prop_missing_values = missing_values
        if hidden_cell_values is not None:
            self.prop_hidden_cell_values = hidden_cell_values
        if hide_legend is not None:
            self.prop_hide_legend = hide_legend

    # region Overrides
    def _supported_services(self) -> Tuple[str, ...]:
        try:
            return self._supported_services_values
        except AttributeError:
            self._supported_services_values = ()
        return self._supported_services_values

    def _is_valid_obj(self, obj: object) -> bool:
        return self._chart_doc is not None

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
        # diagram_provider = mLo.Lo.qi(XDiagramProvider, self._chart_doc)
        # diagram = cast("Union[Diagram, DataProvider]", mProps.Props.get(self._chart_doc, "Diagram"))
        diagram = self._chart_doc.getDiagram()
        diagram_props = {}
        if self.prop_side_by_side is not None:
            diagram_props[self._props.side_by_side] = self.prop_side_by_side

        if self.prop_hidden_cell_values is not None:
            diagram_props[self._props.hidden_cell_values] = self.prop_hidden_cell_values

        if self.prop_missing_values is not None:
            diagram_props[self._props.missing_values] = self.prop_missing_values.value

        if self.prop_primary_y_axis is not None:
            # when primary_y_axis is changed, we need to toggle has_y_axis and has_secondary_y_axis
            diagram_props[self._props.has_y_axis] = self._get(self._props.has_y_axis)
            diagram_props[self._props.has_secondary_y_axis] = self._get(self._props.has_secondary_y_axis)

        if len(diagram_props) > 0:
            mProps.Props.set(diagram, **diagram_props)

        ds = Chart2.get_data_series(chart_doc=self._chart_doc)
        ds1 = ds[0]
        ds_props = {}
        if self.prop_primary_y_axis is not None:
            if self.prop_primary_y_axis:
                ds_props[self._props.primary_y_axis] = 0
            else:
                ds_props[self._props.primary_y_axis] = 1
        if self.prop_hide_legend is not None:
            ds_props[self._props.show_legend] = self.prop_hide_legend

        if len(ds_props) > 0:
            mProps.Props.set(ds1, **ds_props)

        spacing = self.prop_spacing
        overlap = self.prop_overlap
        if spacing is not None or overlap is not None:
            # get the axis to apply spacing and overlap.
            # The axis is determined by the primary_y_axis property of the data series.
            # For this reason, the data series primary_y_axis property must be set first.
            # if primary_y_axis is True, then the YAxis is used, otherwise the SecondaryYAxis is used.
            is_primary_y = int(mProps.Props.get(ds1, self._props.primary_y_axis, 0)) == 0
            if is_primary_y:
                axis = cast(XPropertySet, mProps.Props.get(diagram, "YAxis"))
            else:
                axis = cast(XPropertySet, mProps.Props.get(diagram, "SecondaryYAxis"))

            axis_props = {}
            if spacing is not None:
                axis_props[self._props.spacing] = spacing
            if overlap is not None:
                axis_props[self._props.overlap] = overlap
            mProps.Props.set(axis, **axis_props)

    # endregion apply()

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
        return self._get(self._props.primary_y_axis)

    @prop_primary_y_axis.setter
    def prop_primary_y_axis(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.primary_y_axis)
            self._remove(self._props.has_y_axis)
            self._remove(self._props.has_secondary_y_axis)
            return
        self._set(self._props.primary_y_axis, value)
        self._set(self._props.has_y_axis, value)
        self._set(self._props.has_secondary_y_axis, not value)

    @property
    def prop_spacing(self) -> int | None:
        """Gets or sets the spacing between bars."""
        return self._get(self._props.spacing)

    @prop_spacing.setter
    def prop_spacing(self, value: int | None) -> None:
        if value is None:
            self._remove(self._props.spacing)
            return
        if value < 0:
            value = 0
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
        return self._get(self._props.side_by_side)

    @prop_side_by_side.setter
    def prop_side_by_side(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.side_by_side)
            return
        self._set(self._props.side_by_side, value)

    @property
    def prop_missing_values(self) -> MissingValueTreatmentEnum | None:
        """Gets or sets the missing value treatment."""
        pv = cast(int, self._get(self._props.missing_values))
        if pv is None:
            return None
        return MissingValueTreatmentEnum(pv)

    @prop_missing_values.setter
    def prop_missing_values(self, value: MissingValueTreatmentEnum | None) -> None:
        if value is None:
            self._remove(self._props.missing_values)
            return
        self._set(self._props.missing_values, value.value)

    @property
    def prop_hidden_cell_values(self) -> bool | None:
        """Gets or sets whether hidden cells are included."""
        return self._get(self._props.hidden_cell_values)

    @prop_hidden_cell_values.setter
    def prop_hidden_cell_values(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.hidden_cell_values)
            return
        self._set(self._props.hidden_cell_values, value)

    @property
    def prop_hide_legend(self) -> bool | None:
        """Gets or sets whether the legend is hidden."""
        return self._get(self._props.show_legend)

    @prop_hide_legend.setter
    def prop_hide_legend(self, value: bool | None) -> None:
        if value is None:
            self._remove(self._props.show_legend)
            return
        self._set(self._props.show_legend, not value)

    @property
    def _props(self) -> _OptionsProps:
        try:
            return self._props_internal_attributes
        except AttributeError:
            self._props_internal_attributes = _OptionsProps(
                primary_y_axis="AttachedAxisIndex",
                spacing="GapWidth",
                overlap="Overlap",
                side_by_side="GroupBarsPerAxis",
                missing_values="MissingValueTreatment",
                hidden_cell_values="IncludeHiddenCells",
                show_legend="ShowLegendEntry",
                has_y_axis="HasYAxis",
                has_secondary_y_axis="HasSecondaryYAxis",
            )
        return self._props_internal_attributes

    # endregion Properties
