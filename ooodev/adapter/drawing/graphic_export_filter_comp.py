from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase

from ooodev.adapter.drawing.graphic_export_filter_partial import GraphicExportFilterPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import GraphicExportFilter  # service


class GraphicExportFilterComp(ComponentBase, GraphicExportFilterPartial):
    """
    Class for managing table GraphicExportFilter Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: GraphicExportFilter) -> None:
        """
        Constructor

        Args:
            component (GraphicExportFilter): UNO GraphicExportFilter Component.
        """
        ComponentBase.__init__(self, component)
        GraphicExportFilterPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.GraphicExportFilter",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> GraphicExportFilter:
        """GraphicExportFilter Component"""
        # pylint: disable=no-member
        return cast("GraphicExportFilter", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
