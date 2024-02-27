from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.drawing.draw_page_partial import DrawPagePartial
from ooodev.adapter.drawing.shape_grouper_partial import ShapeGrouperPartial
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.sheet import SpreadsheetDrawPage  # service

# Draw page does implement XDrawPage, but it show in the API of DrawPage Service.


class SpreadsheetDrawPageComp(ComponentBase, DrawPagePartial, ShapeGrouperPartial):
    """
    Class for managing table SpreadsheetDrawPage Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.sheet.SpreadsheetDrawPage`` service.
        """
        ComponentBase.__init__(self, component)
        DrawPagePartial.__init__(self, component, None)
        ShapeGrouperPartial.__init__(self, component, None)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SpreadsheetDrawPage",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> SpreadsheetDrawPage:
        """SpreadsheetDrawPage Component"""
        # pylint: disable=no-member
        return cast("SpreadsheetDrawPage", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
