# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import
from ooodev.adapter.awt.uno_control_fixed_hyperlink_model_partial import UnoControlFixedHyperlinkModelPartial
from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from .ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedHyperlink  # service
    from com.sun.star.awt import UnoControlFixedHyperlinkModel  # service
# endregion imports


class CtlHyperlinkFixed(DialogControlBase, UnoControlFixedHyperlinkModelPartial, ActionEvents):
    """Class for Fixed Hyperlink Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlFixedHyperlink) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlFixedHyperlink): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlFixedHyperlinkModelPartial.__init__(self, self.get_model())
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ActionEvents.__init__(self, trigger_args=generic_args, cb=self._on_action_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_action_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addActionListener(self.events_listener_action)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlFixedHyperlink:
        return cast("UnoControlFixedHyperlink", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlFixedHyperlink``"""
        return "com.sun.star.awt.UnoControlFixedHyperlink"

    def get_model(self) -> UnoControlFixedHyperlinkModel:
        """Gets the Model for the control"""
        return cast("UnoControlFixedHyperlinkModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.HYPERLINK``"""
        return DialogControlKind.HYPERLINK

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.HYPERLINK``"""
        return DialogControlNamedKind.HYPERLINK

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlFixedHyperlink:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlFixedHyperlinkModel:
        return self.get_model()

    # endregion Properties
