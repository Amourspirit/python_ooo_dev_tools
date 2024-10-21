from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.style.style_comp import StyleComp
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
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertiesChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Overrides
    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.style.CellStyle",)

    # endregion Overrides
    # region Properties

    @property
    @override
    def component(self) -> CellStyle:
        """CellStyle Component"""
        # pylint: disable=no-member
        return cast("CellStyle", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
