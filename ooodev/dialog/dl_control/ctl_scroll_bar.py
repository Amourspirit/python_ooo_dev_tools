# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.adjustment_events import AdjustmentEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind

from .ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlScrollBar  # service
    from com.sun.star.awt import UnoControlScrollBarModel  # service
# endregion imports


class CtlScrollBar(DialogControlBase, AdjustmentEvents):
    """Class for Scroll Bar Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlScrollBar) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlScrollBar): Scroll Bar Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        AdjustmentEvents.__init__(self, trigger_args=generic_args, cb=self._on_adjustment_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_adjustment_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addAdjustmentListener(self.events_listener_adjustment)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlScrollBar:
        return cast("UnoControlScrollBar", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlScrollBar``"""
        return "com.sun.star.awt.UnoControlScrollBar"

    def get_model(self) -> UnoControlScrollBarModel:
        """Gets the Model for the control"""
        return cast("UnoControlScrollBarModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.SCROLL_BAR``"""
        return DialogControlKind.SCROLL_BAR

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.SCROLL_BAR``"""
        return DialogControlNamedKind.SCROLL_BAR

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlScrollBar:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlScrollBarModel:
        return self.get_model()

    @property
    def value(self) -> int:
        """Gets or sets the current value of the scroll bar"""
        return self.view.getValue()

    @value.setter
    def value(self, value: int) -> None:
        self.view.setValue(value)

    # endregion Properties
