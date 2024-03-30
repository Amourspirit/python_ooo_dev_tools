from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.uno_control_edit_comp import UnoControlEditComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFormattedField


class UnoControlFormattedFieldComp(UnoControlEditComp):

    def __init__(self, component: UnoControlFormattedField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlFormattedField`` service.
        """
        UnoControlEditComp.__init__(self, component=component)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlFormattedField",)

    @property
    def component(self) -> UnoControlFormattedField:
        """UnoControlFormattedField Component"""
        # pylint: disable=no-member
        return cast("UnoControlFormattedField", self._ComponentBase__get_component())  # type: ignore
