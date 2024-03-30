from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.fixed_hyperlink_partial import FixedHyperlinkPartial
from ooodev.adapter.awt.layout_constrains_partial import LayoutConstrainsPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedHyperlink


class UnoControlFixedHyperlinkComp(UnoControlComp, FixedHyperlinkPartial, LayoutConstrainsPartial):

    def __init__(self, component: UnoControlFixedHyperlink):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlFixedHyperlink`` service.
        """
        UnoControlComp.__init__(self, component=component)
        FixedHyperlinkPartial.__init__(self, component=self.component, interface=None)
        LayoutConstrainsPartial.__init__(self, component=self.component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlFixedHyperlink",)

    @property
    def component(self) -> UnoControlFixedHyperlink:
        """UnoControlFixedHyperlink Component"""
        # pylint: disable=no-member
        return cast("UnoControlFixedHyperlink", self._ComponentBase__get_component())  # type: ignore
