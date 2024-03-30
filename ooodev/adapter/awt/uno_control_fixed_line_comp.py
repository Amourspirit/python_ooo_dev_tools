from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_comp import UnoControlComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedLine


class UnoControlFixedLineComp(UnoControlComp):

    def __init__(self, component: UnoControlFixedLine):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlFixedLine`` service.
        """
        UnoControlComp.__init__(self, component=component)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlFixedLine",)

    @property
    def component(self) -> UnoControlFixedLine:
        """UnoControlFixedLine Component"""
        # pylint: disable=no-member
        return cast("UnoControlFixedLine", self._ComponentBase__get_component())  # type: ignore
