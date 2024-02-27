from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.drawing.draw_pages_partial import DrawPagesPartial


if TYPE_CHECKING:
    from com.sun.star.drawing import DrawPages  # service

# Draw page does implement XDrawPage, but it show in the API of DrawPages Service.


class DrawPagesComp(ComponentBase, DrawPagesPartial):
    """
    Class for managing table DrawPages Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.drawing.DrawPages`` service.
        """
        ComponentBase.__init__(self, component)
        DrawPagesPartial.__init__(self, component=component, interface=None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.DrawPages",)

    # endregion Overrides
    # region Properties

    @property
    def component(self) -> DrawPages:
        """DrawPages Component"""
        # pylint: disable=no-member
        return cast("DrawPages", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
