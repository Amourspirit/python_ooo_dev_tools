from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.spin_value_partial import SpinValuePartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlSpinButton


class UnoControlSpinButtonComp(UnoControlComp, SpinValuePartial):

    def __init__(self, component: UnoControlSpinButton):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlSpinButton`` service.
        """
        UnoControlComp.__init__(self, component=component)
        SpinValuePartial.__init__(self, component=self.component, interface=None)

    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlSpinButton",)

    @property
    @override
    def component(self) -> UnoControlSpinButton:
        """UnoControlSpinButton Component"""
        # pylint: disable=no-member
        return cast("UnoControlSpinButton", self._ComponentBase__get_component())  # type: ignore
