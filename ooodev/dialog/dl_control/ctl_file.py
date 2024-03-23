# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import uno  # pylint: disable=unused-import

from ooodev.mock import mock_g
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.adapter.awt.uno_control_file_control_model_partial import UnoControlFileControlModelPartial
from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFileControl  # service
    from com.sun.star.awt import UnoControlFileControlModel  # service
    from ooodev.dialog.dl_control.model.model_file import ModelFile
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
        UnoControlFileControlModelPartial.__init__(self, component=self.get_model())
        generic_args = self._get_generic_args()
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)
        self._model_ex = None

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
    def model_ex(self) -> ModelFile:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_file import ModelFile

            self._model_ex = ModelFile(self.model)
        return self._model_ex

    @property
    def view(self) -> UnoControlFileControl:
        # pylint: disable=no-member
        return cast("UnoControlFileControl", super().view)

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_file import ModelFile
