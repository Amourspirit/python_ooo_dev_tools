from __future__ import annotations
from typing import cast, TYPE_CHECKING
import contextlib
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.chart2.diagram_partial import DiagramPartial
from ooodev.adapter.chart2.coordinate_system_container_partial import CoordinateSystemContainerPartial
from ooodev.adapter.chart2.titled_partial import TitledPartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import Diagram  # service
    from com.sun.star.chart2 import RelativePosition  # Struct
    from com.sun.star.chart2 import RelativeSize  # Struct


class DiagramComp(ComponentBase, DiagramPartial, CoordinateSystemContainerPartial, TitledPartial):
    """
    Class for managing Chart2 Diagram Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Diagram) -> None:
        """
        Constructor

        Args:
            component (Diagram): UNO component that implements ``com.sun.star.chart2.Diagram`` service.
        """
        ComponentBase.__init__(self, component)
        DiagramPartial.__init__(self, component=component, interface=None)
        CoordinateSystemContainerPartial.__init__(self, component=component, interface=None)
        TitledPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.chart2.Diagram",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> Diagram:
        """Diagram Component"""
        # pylint: disable=no-member
        return cast("Diagram", self._ComponentBase__get_component())  # type: ignore

    @property
    def connect_bars(self) -> bool | None:
        """
        Gets/Sets - Draw connection lines for stacked bar charts.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.ConnectBars
        return None

    @connect_bars.setter
    def connect_bars(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.ConnectBars = value

    @property
    def data_table_h_border(self) -> bool | None:
        """
        Gets/Sets - Chart Data table flags.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.DataTableHBorder
        return None

    @data_table_h_border.setter
    def data_table_h_border(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.DataTableHBorder = value

    @property
    def data_table_outline(self) -> bool | None:
        """
        Gets/Sets the Data Table Outline property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.DataTableOutline
        return None

    @data_table_outline.setter
    def data_table_outline(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.DataTableOutline = value

    @property
    def data_table_v_border(self) -> bool | None:
        """
        Gets/Sets the Data Table V Border property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.DataTableVBorder
        return None

    @data_table_v_border.setter
    def data_table_v_border(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.DataTableVBorder = value

    @property
    def external_data(self) -> str | None:
        """
        Gets/Sets the External Data property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.ExternalData
        return None

    @external_data.setter
    def external_data(self, value: str) -> None:
        with contextlib.suppress(AttributeError):
            self.component.ExternalData = value

    @property
    def group_bars_per_axis(self) -> bool | None:
        """
        Gets/Sets if bars of a bar or column chart are attached to different axis, this property determines how to display those.

        If ``True``, the bars are grouped together in one block for each axis, thus they are painted one group over the other.

        If ``False``, the bars are displayed side-by-side, as if they were all attached to the same axis.

        If all data series of a bar or column chart are attached to only one axis, this property has no effect.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.GroupBarsPerAxis
        return None

    @group_bars_per_axis.setter
    def group_bars_per_axis(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.GroupBarsPerAxis = value

    @property
    def missing_value_treatment(self) -> int | None:
        """
        Gets/Sets how empty or invalid cells in the provided data should be handled when displayed.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.MissingValueTreatment
        return None

    @missing_value_treatment.setter
    def missing_value_treatment(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.component.MissingValueTreatment = value

    @property
    def perspective(self) -> int | None:
        """
        Gets/Sets perspective of 3D charts ( [0,100] ).

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.Perspective
        return None

    @perspective.setter
    def perspective(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.component.Perspective = value

    @property
    def pos_size_exclude_labels(self) -> bool | None:
        """
        The attributes RelativePosition and RelativeSize should be used for the inner coordinate region without axis labels and without data labels.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.PosSizeExcludeLabels
        return None

    @pos_size_exclude_labels.setter
    def pos_size_exclude_labels(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.PosSizeExcludeLabels = value

    @property
    def relative_position(self) -> RelativePosition | None:
        """
        The position is as a relative position on the page.

        If a relative position is given the diagram is not automatically placed, but instead is placed relative on the page.

        **May be None**
        """
        return self.component.RelativePosition

    @relative_position.setter
    def relative_position(self, value: RelativePosition) -> None:
        self.component.RelativePosition = value

    @property
    def relative_size(self) -> RelativeSize:
        """
        The size of the diagram as relative size of the page size.
        """
        return self.component.RelativeSize

    @relative_size.setter
    def relative_size(self, value: RelativeSize) -> None:
        self.component.RelativeSize = value

    @property
    def right_angled_axes(self) -> bool | None:
        """
        Gets/Sets Right Angled Axes property.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.RightAngledAxes
        return None

    @right_angled_axes.setter
    def right_angled_axes(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.RightAngledAxes = value

    @property
    def rotation_horizontal(self) -> int | None:
        """
        Gets/Sets horizontal rotation of 3D charts in degrees ( [-180,180] ).

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.RotationHorizontal
        return None

    @rotation_horizontal.setter
    def rotation_horizontal(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.component.RotationHorizontal = value

    @property
    def rotation_vertical(self) -> int | None:
        """
        Gets/Sets vertical rotation of 3D charts in degrees ( ]-180,180] ).

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.RotationVertical
        return None

    @rotation_vertical.setter
    def rotation_vertical(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.component.RotationVertical = value

    @property
    def sort_by_x_values(self) -> bool | None:
        """
        Gets/Sets - Sort data points by x values for rendering.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.SortByXValues
        return None

    @sort_by_x_values.setter
    def sort_by_x_values(self, value: bool) -> None:
        with contextlib.suppress(AttributeError):
            self.component.SortByXValues = value

    @property
    def starting_angle(self) -> int | None:
        """
        Gets/Sets - Starting angle in degrees for pie charts and doughnut charts.

        **optional**
        """
        with contextlib.suppress(AttributeError):
            return self.component.StartingAngle
        return None

    @starting_angle.setter
    def starting_angle(self, value: int) -> None:
        with contextlib.suppress(AttributeError):
            self.component.StartingAngle = value

    # endregion Properties
