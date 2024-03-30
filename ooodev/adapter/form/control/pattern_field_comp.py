from __future__ import annotations
from typing import cast, TYPE_CHECKING
from ooodev.adapter.awt.uno_control_pattern_field_comp import UnoControlPatternFieldComp
from ooodev.adapter.form.bound_control_partial import BoundControlPartial

if TYPE_CHECKING:
    from com.sun.star.form.control import PatternField


class PatternFieldComp(UnoControlPatternFieldComp, BoundControlPartial):
    """Class for PatternField Control"""

    def __init__(self, component: PatternField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.form.control.PatternField`` service.
        """
        UnoControlPatternFieldComp.__init__(self, component=component)
        BoundControlPartial.__init__(self, component=component, interface=None)

    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.control.PatternField",)

    @property
    def component(self) -> PatternField:
        """PatternField Component"""
        # pylint: disable=no-member
        return cast("PatternField", self._ComponentBase__get_component())  # type: ignore
