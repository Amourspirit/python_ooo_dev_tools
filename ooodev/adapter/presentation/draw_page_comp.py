from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.drawing.draw_page_comp import DrawPageComp as DrawingDrawPageComp
from ooodev.adapter.document.link_target_comp import LinkTargetComp


if TYPE_CHECKING:
    from com.sun.star.presentation import DrawPage  # service


class DrawPageComp(DrawingDrawPageComp, LinkTargetComp):
    """
    Class for managing table DrawPage Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: DrawPage) -> None:
        """
        Constructor

        Args:
            component (DrawPage): UNO DrawPage Component.
        """
        DrawingDrawPageComp.__init__(self, component)
        LinkTargetComp.__init__(self, component)

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.DrawPage",)

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> DrawPage:
            """DrawPage Component"""
            return cast("DrawPage", self._get_component())

    # endregion Properties
