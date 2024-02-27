from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from com.sun.star.chart2 import XLegend
from ooodev.adapter.beans.properties_change_implement import PropertiesChangeImplement
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.drawing.fill_properties_partial import FillPropertiesPartial
from ooodev.adapter.drawing.line_properties_partial import LinePropertiesPartial
from ooodev.adapter.beans.property_set_partial import PropertySetPartial
from ooodev.adapter.style.character_properties_partial import CharacterPropertiesPartial

if TYPE_CHECKING:
    from com.sun.star.chart2 import Legend  # service
    from com.sun.star.awt import Size  # struct
    from com.sun.star.chart2 import RelativePosition  # Struct
    from com.sun.star.chart2.LegendPosition import LegendPositionProto  # type: ignore
    from com.sun.star.chart.ChartLegendExpansion import ChartLegendExpansionProto  # type: ignore
    from ooodev.loader.inst.lo_inst import LoInst


class LegendComp(
    ComponentBase,
    FillPropertiesPartial,
    LinePropertiesPartial,
    CharacterPropertiesPartial,
    PropertySetPartial,
    PropertiesChangeImplement,
    PropertyChangeImplement,
    VetoableChangeImplement,
):
    """
    Class for managing Chart2 Legend Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, lo_inst: LoInst, component: Legend | None = None) -> None:
        """
        Constructor

        Args:
            lo_inst (LoInst): Lo Instance. This instance is used to create ``component`` is it is not provided.
            component (Legend, optional): UNO Chart2 Legend Component.
        """
        if component is None:
            component = cast(
                "Legend", lo_inst.create_instance_mcf(XLegend, "com.sun.star.chart2.Legend", raise_err=True)
            )
        ComponentBase.__init__(self, component)
        FillPropertiesPartial.__init__(self, component=self.component)
        LinePropertiesPartial.__init__(self, component=self.component)
        CharacterPropertiesPartial.__init__(self, component=self.component)  # type: ignore
        PropertySetPartial.__init__(self, component=self.component, interface=None)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertiesChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.chart2.Legend",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> Legend:
        """Legend Component"""
        # pylint: disable=no-member
        return cast("Legend", self._ComponentBase__get_component())  # type: ignore

    @property
    def anchor_position(self) -> LegendPositionProto:
        """
        Gets/Sets - Provides an automated position.
        """
        return self.component.AnchorPosition

    @anchor_position.setter
    def anchor_position(self, value: LegendPositionProto) -> None:
        self.component.AnchorPosition = value

    @property
    def expansion(self) -> ChartLegendExpansionProto:
        """
        Gets/Sets how the aspect ratio of the legend should roughly be.

        Set the Expansion to ``com.sun.star.chart.HIGH`` for a legend that is positioned on the right or left hand side.
        Use ``com.sun.star.chart.WIDE`` for a legend that is positioned on top or the bottom.
        """
        return self.component.Expansion

    @expansion.setter
    def expansion(self, value: ChartLegendExpansionProto) -> None:
        self.component.Expansion = value

    @property
    def overlay(self) -> bool:
        """
        Gets/Sets whether the legend should overlay the chart.

        **since**

            LibreOffice 7.0
        """
        return self.component.Overlay

    @overlay.setter
    def overlay(self, value: bool) -> None:
        self.component.Overlay = value

    @property
    def reference_page_size(self) -> Size | None:
        """
        Gets/Sets the size of the page at the time when properties were set (e.g. the CharHeight).

        This way it is possible to resize objects (like text) in the view without modifying the model.

        **My be None**
        """
        return self.component.ReferencePageSize

    @reference_page_size.setter
    def reference_page_size(self, value: Size) -> None:
        self.component.ReferencePageSize = value

    @property
    def relative_position(self) -> RelativePosition | None:
        """
        Gets/Sets the position is as a relative position on the page.

        If a relative position is given the legend is not automatically placed, but instead is placed relative on the page.

        If ``None``, the legend position is solely determined by the ``anchor_position``.

        **My be None**
        """
        return self.component.RelativePosition

    @relative_position.setter
    def relative_position(self, value: RelativePosition) -> None:
        self.component.RelativePosition = value

    @property
    def show(self) -> bool:
        """
        Gets/Sets whether the legend should be rendered by the view.
        """
        return self.component.Show

    @show.setter
    def show(self, value: bool) -> None:
        self.component.Show = value

    # endregion Properties
