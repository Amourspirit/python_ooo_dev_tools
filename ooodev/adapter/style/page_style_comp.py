from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.style.style_comp import StyleComp
from ooodev.adapter.beans.properties_change_implement import PropertiesChangeImplement

if TYPE_CHECKING:
    from com.sun.star.style import PageStyle  # service
    from com.sun.star.style import XStyle


class PageStyleComp(StyleComp, PropertiesChangeImplement):
    """
    Class for managing table PageStyle Component.
    """

    def __init__(self, component: XStyle) -> None:
        """
        Constructor

        Args:
            component (XStyle): UNO Component that supports ``com.sun.star.style.PageStyle`` service.
        """
        StyleComp.__init__(self, component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertiesChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.style.PageStyle",)

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> PageStyle:
            """PageStyle Component"""
            # pylint: disable=no-member
            return cast("PageStyle", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
