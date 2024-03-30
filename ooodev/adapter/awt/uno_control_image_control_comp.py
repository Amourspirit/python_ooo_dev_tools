from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.layout_constrains_partial import LayoutConstrainsPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlImageControl


class UnoControlImageControlComp(UnoControlComp, LayoutConstrainsPartial):

    def __init__(self, component: UnoControlImageControl):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlImageControl`` service.
        """
        UnoControlComp.__init__(self, component=component)
        LayoutConstrainsPartial.__init__(self, component=self.component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlImageControl",)

    @property
    def component(self) -> UnoControlImageControl:
        """UnoControlImageControl Component"""
        # pylint: disable=no-member
        return cast("UnoControlImageControl", self._ComponentBase__get_component())  # type: ignore
