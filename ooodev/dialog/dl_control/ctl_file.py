# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.adapter.awt.uno_control_file_control_model_partial import UnoControlFileControlModelPartial
from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFileControl  # service
    from com.sun.star.awt import UnoControlFileControlModel  # service
# endregion imports


class CtlFile(DialogControlBase, UnoControlFileControlModelPartial, TextEvents):
    """Class for file Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlFileControl) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlFileControl): File Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlFileControlModelPartial.__init__(self)
        generic_args = self._get_generic_args()
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTextListener(self.events_listener_text)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlFileControl:
        return cast("UnoControlFileControl", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlFileControl``"""
        return "com.sun.star.awt.UnoControlFileControl"

    def get_model(self) -> UnoControlFileControlModel:
        """Gets the Model for the control"""
        return cast("UnoControlFileControlModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.FILE_CONTROL``"""
        return DialogControlKind.FILE_CONTROL

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.FILE_CONTROL``"""
        return DialogControlNamedKind.FILE_CONTROL

    # endregion Overrides

    # region Properties

    @property
    def model(self) -> UnoControlFileControlModel:
        # pylint: disable=no-member
        return cast("UnoControlFileControlModel", super().model)

    @property
    def view(self) -> UnoControlFileControl:
        # pylint: disable=no-member
        return cast("UnoControlFileControl", super().view)

    # endregion Properties
