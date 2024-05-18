# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.mock import mock_g
from ooodev.adapter.awt.uno_control_radio_button_model_partial import UnoControlRadioButtonModelPartial
from ooodev.adapter.awt.item_events import ItemEvents

# pylint: disable=useless-import-alias
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase, _create_control


if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlRadioButton  # service
    from com.sun.star.awt import UnoControlRadioButtonModel  # service
    from com.sun.star.awt import XWindowPeer
    from ooodev.dialog.dl_control.model.model_radio_button import ModelRadioButton
    from ooodev.dialog.dl_control.view.view_radio_button import ViewRadioButton
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
        UnoControlRadioButtonModelPartial.__init__(self, component=self.get_model())
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ItemEvents.__init__(self, trigger_args=generic_args, cb=self._on_item_event_listener_add_remove)
        self._model_ex = None
        self._view_ex = None

    # endregion init

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlRadioButton({self.name})"
        return "CtlRadioButton"

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

    # region Static Methods
    @staticmethod
    def create(win: XWindowPeer, **kwargs: Any) -> "CtlRadioButton":
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

        Returns:
            CtlRadioButton: New instance of the control.

        Note:
            The `UnoControlDialogElement <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogElement.html>`__
            interface is not included when creating the control with a window peer.
        """
        ctrl = _create_control("com.sun.star.awt.UnoControlRadioButtonModel", win, **kwargs)
        return CtlRadioButton(ctl=ctrl)

    # endregion Static Methods

    # region Properties

    @property
    def model(self) -> UnoControlRadioButtonModel:
        # pylint: disable=no-member
        return cast("UnoControlRadioButtonModel", super().model)

    @property
    def model_ex(self) -> ModelRadioButton:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_radio_button import ModelRadioButton

            self._model_ex = ModelRadioButton(self.model)
        return self._model_ex

    @property
    def view(self) -> UnoControlRadioButton:
        # pylint: disable=no-member
        return cast("UnoControlRadioButton", super().view)

    @property
    def view_ex(self) -> ViewRadioButton:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        # pylint: disable=no-member
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_radio_button import ViewRadioButton

            self._view_ex = ViewRadioButton(self.view)
        return self._view_ex

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_radio_button import ModelRadioButton
    from ooodev.dialog.dl_control.view.view_radio_button import ViewRadioButton
