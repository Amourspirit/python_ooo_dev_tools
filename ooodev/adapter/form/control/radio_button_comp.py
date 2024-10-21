from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.awt.uno_control_radio_button_comp import UnoControlRadioButtonComp
from ooodev.adapter.form.bound_control_partial import BoundControlPartial

if TYPE_CHECKING:
    from com.sun.star.form.control import RadioButton


class RadioButtonComp(UnoControlRadioButtonComp, BoundControlPartial):
    """Class for RadioButton Control"""

    def __init__(self, component: RadioButton):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.form.control.RadioButton`` service.
        """
        UnoControlRadioButtonComp.__init__(self, component=component)
        BoundControlPartial.__init__(self, component=component, interface=None)

    @override
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.control.RadioButton",)

    @property
    @override
    def component(self) -> RadioButton:
        """RadioButton Component"""
        # pylint: disable=no-member
        return cast("RadioButton", self._ComponentBase__get_component())  # type: ignore
