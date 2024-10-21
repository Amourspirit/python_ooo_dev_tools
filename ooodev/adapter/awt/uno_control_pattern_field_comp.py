from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.awt.uno_control_edit_comp import UnoControlEditComp
from ooodev.adapter.awt.pattern_field_partial import PatternFieldPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlPatternField


class UnoControlPatternFieldComp(UnoControlEditComp, PatternFieldPartial):

    def __init__(self, component: UnoControlPatternField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlPatternField`` service.
        """
        UnoControlEditComp.__init__(self, component=component)
        PatternFieldPartial.__init__(self, component=component, interface=None)

    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlPatternField",)

    @property
    @override
    def component(self) -> UnoControlPatternField:
        """UnoControlPatternField Component"""
        # pylint: disable=no-member
        return cast("UnoControlPatternField", self._ComponentBase__get_component())  # type: ignore
