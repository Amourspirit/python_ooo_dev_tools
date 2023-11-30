# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

# pylint: disable=useless-import-alias
from ooodev.adapter.awt.item_events import ItemEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.utils.kind.tri_state_kind import TriStateKind as TriStateKind
from .ctl_base import DialogControlBase


if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlCheckBox  # service
    from com.sun.star.awt import UnoControlCheckBoxModel  # service
# endregion imports


class CtlCheckBox(DialogControlBase, ItemEvents):
    """Class for CheckBox Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlCheckBox) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlCheckBox): Check Box Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ItemEvents.__init__(self, trigger_args=generic_args, cb=self._on_item_event_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_item_event_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addItemListener(self.events_listener_item)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlCheckBox:
        return cast("UnoControlCheckBox", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlCheckBox``"""
        return "com.sun.star.awt.UnoControlCheckBox"

    def get_model(self) -> UnoControlCheckBoxModel:
        """Gets the Model for the control"""
        return cast("UnoControlCheckBoxModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.CHECKBOX``"""
        return DialogControlKind.CHECKBOX

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.CHECKBOX``"""
        return DialogControlNamedKind.CHECKBOX

    # endregion Overrides

    # region Properties
    @property
    def border(self) -> BorderKind:
        """Gets/Sets the border style"""
        return BorderKind(self.model.VisualEffect)

    @border.setter
    def border(self, value: BorderKind) -> None:
        self.model.VisualEffect = value.value

    @property
    def label(self) -> str:
        """Gets/Sets the label (text) for the control"""
        return self.model.Label

    @label.setter
    def label(self, value: str) -> None:
        self.model.Label = value

    @property
    def model(self) -> UnoControlCheckBoxModel:
        return self.get_model()

    @property
    def state(self) -> TriStateKind:
        """Gets/Sets the state"""
        return TriStateKind(self.model.State)

    @state.setter
    def state(self, value: TriStateKind) -> None:
        self.model.State = value.value

    @property
    def triple_state(self) -> bool:
        """
        Gets/Sets the triple state.

        Specifies if the checkbox control may appear dimmed (grayed) or not.
        """
        return self.model.TriState

    @triple_state.setter
    def triple_state(self, value: bool) -> None:
        self.model.TriState = value

    @property
    def view(self) -> UnoControlCheckBox:
        return self.get_view_ctl()

    # endregion Properties
