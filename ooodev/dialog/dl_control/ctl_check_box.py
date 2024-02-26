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
from ooodev.adapter.awt.uno_control_check_box_model_partial import UnoControlCheckBoxModelPartial
from ooodev.dialog.dl_control.ctl_base import DialogControlBase


if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlCheckBox  # service
    from com.sun.star.awt import UnoControlCheckBoxModel  # service
# endregion imports


class CtlCheckBox(DialogControlBase, UnoControlCheckBoxModelPartial, ItemEvents):
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
        UnoControlCheckBoxModelPartial.__init__(self)
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
    def model(self) -> UnoControlCheckBoxModel:
        # pylint: disable=no-member
        return cast("UnoControlCheckBoxModel", super().model)

    @property
    def triple_state(self) -> bool:
        """
        Gets/Sets the triple state Same as ``tri_state`` property.

        Specifies if the checkbox control may appear dimmed (grayed) or not.
        """
        return self.tri_state

    @triple_state.setter
    def triple_state(self, value: bool) -> None:
        self.tri_state = value

    @property
    def view(self) -> UnoControlCheckBox:
        # pylint: disable=no-member
        return cast("UnoControlCheckBox", super().view)

    # endregion Properties
