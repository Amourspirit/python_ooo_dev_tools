from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.button_partial import ButtonPartial
from ooodev.adapter.awt.layout_constrains_partial import LayoutConstrainsPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlButton


class UnoControlButtonComp(UnoControlComp, ButtonPartial, LayoutConstrainsPartial):

    def __init__(self, component: UnoControlButton):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlButton`` service.
        """
        UnoControlComp.__init__(self, component=component)
        ButtonPartial.__init__(self, component=self.component, interface=None)
        LayoutConstrainsPartial.__init__(self, component=self.component, interface=None)

    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlButton",)

    @property
    @override
    def component(self) -> UnoControlButton:
        """UnoControlButton Component"""
        # pylint: disable=no-member
        return cast("UnoControlButton", self._ComponentBase__get_component())  # type: ignore
