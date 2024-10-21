from __future__ import annotations
from typing import cast, Any, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

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
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.presentation.DrawPage",)

    # endregion Overrides
    # region Properties

    @property
    @override
    def component(self) -> DrawPage:
        """DrawPage Component"""
        return cast("DrawPage", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
