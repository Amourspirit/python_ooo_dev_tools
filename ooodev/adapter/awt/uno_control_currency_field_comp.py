from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno
from ooodev.adapter.awt.currency_field_partial import CurrencyFieldPartial
from ooodev.adapter.awt.uno_control_edit_comp import UnoControlEditComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlCurrencyField


class UnoControlCurrencyFieldComp(UnoControlEditComp, CurrencyFieldPartial):

    def __init__(self, component: UnoControlCurrencyField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlCurrencyField`` service.
        """
        UnoControlEditComp.__init__(self, component=component)
        CurrencyFieldPartial.__init__(self, component=self.component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlCurrencyField",)

    @property
    def component(self) -> UnoControlCurrencyField:
        """UnoControlCurrencyField Component"""
        # pylint: disable=no-member
        return cast("UnoControlCurrencyField", self._ComponentBase__get_component())  # type: ignore
