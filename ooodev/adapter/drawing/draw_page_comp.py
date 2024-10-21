from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.drawing.generic_draw_page_comp import GenericDrawPageComp


if TYPE_CHECKING:
    from com.sun.star.drawing import DrawPage  # service

# Draw page does implement XDrawPage, but it show in the API of DrawPage Service.


class DrawPageComp(GenericDrawPageComp):
    """
    Class for managing table DrawPage Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.drawing.DrawPage`` service.
        """
        GenericDrawPageComp.__init__(self, component)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return (
            "com.sun.star.drawing.DrawPage",
            "com.sun.star.drawing.GenericDrawPage",
        )

    # endregion Overrides
    # region Properties

    @property
    @override
    def component(self) -> DrawPage:
        """DrawPage Component"""
        # pylint: disable=no-member
        return cast("DrawPage", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
