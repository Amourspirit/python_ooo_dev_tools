# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import
from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.type_var import PathOrStr
from ooodev.adapter.awt.uno_control_button_model_partial import UnoControlButtonModelPartial
from ooodev.dialog.dl_control.model.model_button import ModelButton


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
        self.__model = None
        self.model: UnoControlButtonModel
        DialogControlBase.__init__(self, ctl)
        UnoControlButtonModelPartial.__init__(self, self.model)
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
    def model_button(self) -> ModelButton:
        """
        Gets the Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self.__model is None:
            self.__model = ModelButton(cast("UnoControlButtonModel", self.model))
        return self.__model

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
        return self.model_button.picture

    @picture.setter
    def picture(self, value: PathOrStr) -> None:
        self.model_button.picture = value

    # endregion Properties
