from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.awt.item_events import ItemEvents
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.tri_state_kind import TriStateKind as TriStateKind
from .ctl_base import CtlBase


if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlCheckBox  # service
    from com.sun.star.awt import UnoControlCheckBoxModel  # service


class CtlCheckBox(CtlBase, ItemEvents):
    def __init__(self, ctl: UnoControlCheckBox) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlCheckBox): Fixed Text Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ItemEvents.__init__(self, trigger_args=generic_args)
        self.view.addItemListener(self.events_listener_item)

    def get_view_ctl(self) -> UnoControlCheckBox:
        return cast("UnoControlCheckBox", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlCheckBox``"""
        return "com.sun.star.awt.UnoControlCheckBox"

    @property
    def view(self) -> UnoControlCheckBox:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlCheckBoxModel:
        return cast("UnoControlCheckBoxModel", self.get_view_ctl().getModel())

    @property
    def border(self) -> BorderKind:
        """Gets/Sets the border style"""
        return BorderKind(self.model.VisualEffect)

    @border.setter
    def border(self, value: BorderKind) -> None:
        self.model.VisualEffect = value.value

    @property
    def state(self) -> TriStateKind:
        """Gets/Sets the state"""
        return TriStateKind(self.model.State)

    @state.setter
    def state(self, value: TriStateKind) -> None:
        self.model.State = value.value
