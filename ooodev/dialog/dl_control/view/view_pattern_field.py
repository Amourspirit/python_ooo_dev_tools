from __future__ import annotations
from typing import TYPE_CHECKING
from ooodev.adapter.awt.uno_control_pattern_field_comp import UnoControlPatternFieldComp
from ooodev.adapter.awt.spin_field_partial import SpinFieldPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlPatternField


class ViewPatternField(UnoControlPatternFieldComp, SpinFieldPartial):
    """View for the Pattern Field control."""

    def __init__(self, component: UnoControlPatternField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlPatternField`` service.
        """
        UnoControlPatternFieldComp.__init__(self, component=component)
        SpinFieldPartial.__init__(self, component=component)
