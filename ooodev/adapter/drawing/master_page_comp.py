from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.drawing.generic_draw_page_comp import GenericDrawPageComp

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore


if TYPE_CHECKING:
    from com.sun.star.drawing import MasterPage  # service

# Master page does implement XDrawPage, but it show in the API of DrawPage Service.


class MasterPageComp(GenericDrawPageComp):
    """
    Class for managing table MasterPage Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.drawing.MasterPage`` service.
        """
        super().__init__(component)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.drawing.MasterPage",)

    # endregion Overrides
    # region Properties

    @property
    @override
    def component(self) -> MasterPage:
        """MasterPage Component"""
        # pylint: disable=no-member
        return cast("MasterPage", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
