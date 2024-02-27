# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
from pathlib import Path
import uno  # pylint: disable=unused-import
from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.file_io import FileIO
from ooodev.utils.type_var import PathOrStr
from ooodev.adapter.awt.uno_control_button_model_partial import UnoControlButtonModelPartial


from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlButton  # service
    from com.sun.star.awt import UnoControlButtonModel  # service
# endregion imports


class CtlButton(DialogControlBase, UnoControlButtonModelPartial, ActionEvents):
    """Class for Button Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlButton) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlButton): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlButtonModelPartial.__init__(self)
        self._generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ActionEvents.__init__(self, trigger_args=self._generic_args, cb=self._on_action_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_action_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addActionListener(self.events_listener_action)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlButton:
        return cast("UnoControlButton", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlButton``"""
        return "com.sun.star.awt.UnoControlButton"

    def get_model(self) -> UnoControlButtonModel:
        """Gets the Model for the control"""
        return cast("UnoControlButtonModel", self.get_view_ctl().getModel())

    # endregion Overrides

    # region Properties
    @property
    def model(self) -> UnoControlButtonModel:
        # pylint: disable=no-member
        return cast("UnoControlButtonModel", super().model)

    @property
    def view(self) -> UnoControlButton:
        # pylint: disable=no-member
        return cast("UnoControlButton", super().view)

    # region UnoControlButtonModelPartial overrides

    @property
    def image_url(self) -> str:
        """
        Gets/Sets a URL to an image to use for the button.
        """
        return self.picture

    @image_url.setter
    def image_url(self, value: str) -> None:
        self.picture = value

    # endregion UnoControlButtonModelPartial overrides

    @property
    def picture(self) -> str:
        """
        Gets/Sets the picture for the control

        When setting the value it can be a string or a Path object.
        If a string is passed it can be a URL or a path to a file.
        Value such as ``file:///path/to/image.png`` and ``/path/to/image.png`` are valid.
        Relative paths are supported.

        Returns:
            str: The picture URL in the format of ``file:///path/to/image.png`` or empty string if no picture is set.
        """
        with contextlib.suppress(Exception):
            return self.model.ImageURL
        return ""

    @picture.setter
    def picture(self, value: PathOrStr) -> None:
        pth_str = str(value)
        if not pth_str:
            self.model.ImageURL = ""
            return
        if isinstance(value, str):
            if value.startswith("file://"):
                self.model.ImageURL = value
                return
            value = Path(value)
        if not FileIO.is_valid_path_or_str(value):
            raise ValueError(f"Invalid path or str: {value}")
        self.model.ImageURL = FileIO.fnm_to_url(value)

    # endregion Properties
