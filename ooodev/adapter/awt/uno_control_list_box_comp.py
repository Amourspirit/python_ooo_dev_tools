from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_comp import UnoControlComp
from ooodev.adapter.awt.list_box_partial import ListBoxPartial
from ooodev.adapter.awt.layout_constrains_partial import LayoutConstrainsPartial
from ooodev.adapter.awt.text_layout_constrains_partial import TextLayoutConstrainsPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlListBox


class UnoControlListBoxComp(UnoControlComp, ListBoxPartial, LayoutConstrainsPartial, TextLayoutConstrainsPartial):

    def __init__(self, component: UnoControlListBox):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlListBox`` service.
        """
        UnoControlComp.__init__(self, component=component)
        ListBoxPartial.__init__(self, component=self.component, interface=None)
        LayoutConstrainsPartial.__init__(self, component=self.component, interface=None)
        TextLayoutConstrainsPartial.__init__(self, component=self.component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlListBox",)

    @property
    def component(self) -> UnoControlListBox:
        """UnoControlListBox Component"""
        # pylint: disable=no-member
        return cast("UnoControlListBox", self._ComponentBase__get_component())  # type: ignore
