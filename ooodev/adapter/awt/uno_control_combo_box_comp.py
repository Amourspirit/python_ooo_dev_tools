from __future__ import annotations
from typing import cast, TYPE_CHECKING

try:
    # python 3.12+
    from typing import override  # noqa # type: ignore
except ImportError:
    from typing_extensions import override  # noqa # type: ignore

from ooodev.adapter.awt.combo_box_partial import ComboBoxPartial
from ooodev.adapter.awt.uno_control_edit_comp import UnoControlEditComp

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlComboBox


class UnoControlComboBoxComp(UnoControlEditComp, ComboBoxPartial):

    def __init__(self, component: UnoControlComboBox):
        """
        Constructor

        Args:
            component (Any): Component that implements ``com.sun.star.awt.UnoControlComboBox`` service.
        """
        UnoControlEditComp.__init__(self, component=component)
        ComboBoxPartial.__init__(self, component=self.component, interface=None)

    @property
    @override
    def component(self) -> UnoControlComboBox:
        """UnoControlComboBox Component"""
        # pylint: disable=no-member
        return cast("UnoControlComboBox", self._ComponentBase__get_component())  # type: ignore
