# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import
from ooo.dyn.awt.push_button_type import PushButtonType
from ooodev.mock import mock_g
from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.type_var import PathOrStr
from ooodev.adapter.awt.uno_control_button_model_partial import UnoControlButtonModelPartial
from ooodev.dialog.dl_control.ctl_base import DialogControlBase, _create_control

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlButton  # service
    from com.sun.star.awt import UnoControlButtonModel  # service
    from com.sun.star.awt import XWindowPeer
    from ooodev.dialog.dl_control.model.model_button import ModelButton
    from ooodev.dialog.dl_control.view.view_button import ViewButton
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
        self._model_ex = None
        self._view_ex = None
        self.model: UnoControlButtonModel
        DialogControlBase.__init__(self, ctl)
        UnoControlButtonModelPartial.__init__(self, self.model)
        self._generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ActionEvents.__init__(self, trigger_args=self._generic_args, cb=self._on_action_events_listener_add_remove)

    # endregion init

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlButton({self.name})"
        return "CtlButton"

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

    # region Static Methods
    @staticmethod
    def create(win: XWindowPeer, **kwargs: Any) -> "CtlButton":
        """
        Creates a new instance of the control.

        Keyword arguments are optional.
        Extra Keyword args are passed to the control as property values.

        Args:
            win (XWindowPeer): Parent Window

        Keyword Args:
            x (int, UnitT, optional): X Position in Pixels or UnitT.
            y (int, UnitT, optional): Y Position in Pixels or UnitT.
            width (int, UnitT, optional): Width in Pixels or UnitT.
            height (int, UnitT, optional): Height in Pixels or UnitT.
            btn_type (PushButtonType | None, optional): Type of Button.

        Returns:
            CtlButton: New instance of the control.

        Note:
            The `UnoControlDialogElement <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogElement.html>`__
            interface is not included when creating the control with a window peer.
        """
        btn_type = kwargs.pop("btn_type", PushButtonType.STANDARD)
        ctrl = _create_control("com.sun.star.awt.UnoControlButtonModel", win, **kwargs)
        model = ctrl.getModel()
        uno_any = uno.Any("short", btn_type)  # type: ignore
        uno.invoke(model, "setPropertyValue", ("PushButtonType", uno_any))  # type: ignore
        return CtlButton(ctl=ctrl)

    # region Properties
    @property
    def model_ex(self) -> ModelButton:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_button import ModelButton

            self._model_ex = ModelButton(cast("UnoControlButtonModel", self.model))
        return self._model_ex

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
        return self.model_ex.picture

    @picture.setter
    def picture(self, value: PathOrStr) -> None:
        self.model_ex.picture = value

    @property
    def view_ex(self) -> ViewButton:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        # pylint: disable=no-member
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_button import ViewButton

            self._view_ex = ViewButton(self.view)
        return self._view_ex

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_button import ModelButton
    from ooodev.dialog.dl_control.view.view_button import ViewButton
