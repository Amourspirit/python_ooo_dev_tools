# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.uno_control_radio_button_model_partial import UnoControlRadioButtonModelPartial
from ooodev.adapter.awt.item_events import ItemEvents

# pylint: disable=useless-import-alias
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind

from ooodev.dialog.dl_control.ctl_base import DialogControlBase


if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlRadioButton  # service
    from com.sun.star.awt import UnoControlRadioButtonModel  # service
# endregion imports


class CtlRadioButton(DialogControlBase, UnoControlRadioButtonModelPartial, ItemEvents):
    """Class for Radio Button Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlRadioButton) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlRadioButton): Radio Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlRadioButtonModelPartial.__init__(self)
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
    def get_view_ctl(self) -> UnoControlRadioButton:
        return cast("UnoControlRadioButton", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlRadioButton``"""
        return "com.sun.star.awt.UnoControlRadioButton"

    def get_model(self) -> UnoControlRadioButtonModel:
        """Gets the Model for the control"""
        return cast("UnoControlRadioButtonModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.RADIO_BUTTON``"""
        return DialogControlKind.RADIO_BUTTON

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.RADIO_BUTTON``"""
        return DialogControlNamedKind.RADIO_BUTTON

    # endregion Overrides

    # region Properties

    @property
    def model(self) -> UnoControlRadioButtonModel:
        # pylint: disable=no-member
        return cast("UnoControlRadioButtonModel", super().model)

    @property
    def view(self) -> UnoControlRadioButton:
        # pylint: disable=no-member
        return cast("UnoControlRadioButton", super().view)

    # endregion Properties
