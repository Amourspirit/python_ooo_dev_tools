from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.adapter.awt.uno_control_currency_field_comp import UnoControlCurrencyFieldComp
from ooodev.adapter.awt.spin_field_partial import SpinFieldPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlCurrencyField


class ViewCurrencyField(UnoControlCurrencyFieldComp, SpinFieldPartial):
    """View for the Currency Field control."""

    def __init__(self, component: UnoControlCurrencyField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlCurrencyField`` service.
        """
        UnoControlCurrencyFieldComp.__init__(self, component=component)
        SpinFieldPartial.__init__(self, component=component)
