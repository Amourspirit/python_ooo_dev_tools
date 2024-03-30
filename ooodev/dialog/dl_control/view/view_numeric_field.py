from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.adapter.awt.uno_control_numeric_field_comp import UnoControlNumericFieldComp
from ooodev.adapter.awt.spin_field_partial import SpinFieldPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlNumericField


class ViewNumericField(UnoControlNumericFieldComp, SpinFieldPartial):
    """View for the Numeric Field control."""

    def __init__(self, component: UnoControlNumericField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlNumericField`` service.
        """
        UnoControlNumericFieldComp.__init__(self, component=component)
        SpinFieldPartial.__init__(self, component=component)
