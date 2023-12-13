from __future__ import annotations
from typing import cast, Any, TYPE_CHECKING
from ooodev.adapter.drawing.draw_page_comp import DrawPageComp as DrawingDrawPageComp


if TYPE_CHECKING:
    from com.sun.star.presentation import DrawPage  # service


class DrawPageComp(DrawingDrawPageComp):
    """
    Class for managing table DrawPage Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO component that supports ``com.sun.star.presentation.DrawPage`` service.
        """
        DrawingDrawPageComp.__init__(self, component)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.presentation.DrawPage",)

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> DrawPage:
            """DrawPage Component"""
            return cast("DrawPage", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
