from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.fixed_text_partial import FixedTextPartial
from ooodev.adapter.awt.layout_constrains_partial import LayoutConstrainsPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedText


class UnoControlFixedTextComp(UnoControlComp, FixedTextPartial, LayoutConstrainsPartial):

    def __init__(self, component: UnoControlFixedText):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlFixedText`` service.
        """
        UnoControlComp.__init__(self, component=component)
        FixedTextPartial.__init__(self, component=self.component, interface=None)
        LayoutConstrainsPartial.__init__(self, component=self.component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlFixedText",)

    @property
    def component(self) -> UnoControlFixedText:
        """UnoControlFixedText Component"""
        # pylint: disable=no-member
        return cast("UnoControlFixedText", self._ComponentBase__get_component())  # type: ignore
