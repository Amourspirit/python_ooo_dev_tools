# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.uno_control_spin_button_model_partial import UnoControlSpinButtonModelPartial
from ooodev.adapter.awt.spin_value_partial import SpinValuePartial
from ooodev.adapter.awt.adjustment_events import AdjustmentEvents
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlSpinButton  # service
    from com.sun.star.awt import UnoControlSpinButtonModel  # service
# endregion imports


class CtlSpinButton(DialogControlBase, UnoControlSpinButtonModelPartial, SpinValuePartial, AdjustmentEvents):
    """Class for Spin Button Control"""

    # region init
    def __init__(self, ctl: UnoControlSpinButton) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlSpinButton): Progress Bar Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlSpinButtonModelPartial.__init__(self)
        SpinValuePartial.__init__(self, component=self.get_view())  # type: ignore
        generic_args = self._get_generic_args()
        AdjustmentEvents.__init__(self, trigger_args=generic_args, cb=self._on_adjustment_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_adjustment_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addAdjustmentListener(self.events_listener_adjustment)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlSpinButton:
        return cast("UnoControlSpinButton", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlSpinButton``"""
        return "com.sun.star.awt.UnoControlSpinButton"

    def get_model(self) -> UnoControlSpinButtonModel:
        """Gets the Model for the control"""
        return cast("UnoControlSpinButtonModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.SPIN_BUTTON``"""
        return DialogControlKind.SPIN_BUTTON

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.SPIN_BUTTON``"""
        return DialogControlNamedKind.SPIN_BUTTON

    # endregion Overrides

    # region Properties

    @property
    def model(self) -> UnoControlSpinButtonModel:
        # pylint: disable=no-member
        return cast("UnoControlSpinButtonModel", super().model)

    @property
    def view(self) -> UnoControlSpinButton:
        # pylint: disable=no-member
        return cast("UnoControlSpinButton", super().view)

    # endregion Properties
