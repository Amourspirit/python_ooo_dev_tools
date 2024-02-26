from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.drawing.draw_pages_partial import DrawPagesPartial


if TYPE_CHECKING:
    from com.sun.star.drawing import MasterPages  # service


class MasterPagesComp(ComponentBase, DrawPagesPartial):
    """
    Class for managing table MasterPages Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.drawing.MasterPages`` service.
        """
        ComponentBase.__init__(self, component)
        DrawPagesPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.MasterPages",)

    # endregion Overrides
    # region Properties

    @property
    def component(self) -> MasterPages:
        """MasterPages Component"""
        # pylint: disable=no-member
        return cast("MasterPages", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
