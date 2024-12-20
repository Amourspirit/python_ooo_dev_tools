from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.awt.uno_control_edit_comp import UnoControlEditComp
from ooodev.adapter.awt.numeric_field_partial import NumericFieldPartial

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlNumericField


class UnoControlNumericFieldComp(UnoControlEditComp, NumericFieldPartial):

    def __init__(self, component: UnoControlNumericField):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlNumericField`` service.
        """
        UnoControlEditComp.__init__(self, component=component)
        NumericFieldPartial.__init__(self, component=component, interface=None)

    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.awt.UnoControlNumericField",)

    @property
    @override
    def component(self) -> UnoControlNumericField:
        """UnoControlNumericField Component"""
        # pylint: disable=no-member
        return cast("UnoControlNumericField", self._ComponentBase__get_component())  # type: ignore
