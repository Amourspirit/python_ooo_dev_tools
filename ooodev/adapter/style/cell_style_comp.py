from __future__ import annotations
from typing import cast, TYPE_CHECKING
from .style_comp import StyleComp

from ooodev.adapter.beans.properties_change_implement import PropertiesChangeImplement

if TYPE_CHECKING:
    from com.sun.star.style import CellStyle  # service
    from com.sun.star.style import XStyle


class CellStyleComp(StyleComp, PropertiesChangeImplement):
    """
    Class for managing table CellStyle Component.
    """

    def __init__(self, component: XStyle) -> None:
        """
        Constructor

        Args:
            component (XStyle): UNO Component that supports ``com.sun.star.style.CellStyle`` service.
        """
        StyleComp.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertiesChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.style.CellStyle",)

    # endregion Overrides
    # region Properties
    if TYPE_CHECKING:

        @property
        def component(self) -> CellStyle:
            """CellStyle Component"""
            return cast("CellStyle", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
