# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.mock import mock_g
from ooodev.adapter.awt.uno_control_fixed_hyperlink_model_partial import UnoControlFixedHyperlinkModelPartial
from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase, _create_control

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedHyperlink  # service
    from com.sun.star.awt import UnoControlFixedHyperlinkModel  # service
    from com.sun.star.awt import XWindowPeer
    from ooodev.dialog.dl_control.model.model_hyperlink_fixed import ModelHyperlinkFixed
    from ooodev.dialog.dl_control.view.view_fixed_hyperlink import ViewFixedHyperlink
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
        UnoControlFixedHyperlinkModelPartial.__init__(self, component=self.get_model())
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ActionEvents.__init__(self, trigger_args=generic_args, cb=self._on_action_events_listener_add_remove)
        self._model_ex = None
        self._view_ex = None

    # endregion init

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlHyperlinkFixed({self.name})"
        return "CtlHyperlinkFixed"

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

    # region Static Methods
    @staticmethod
    def create(win: XWindowPeer, **kwargs: Any) -> "CtlHyperlinkFixed":
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
            CtlHyperlinkFixed: New instance of the control.

        Note:
            The `UnoControlDialogElement <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogElement.html>`__
            interface is not included when creating the control with a window peer.
        """
        ctrl = _create_control("com.sun.star.awt.UnoControlFixedHyperlinkModel", win, **kwargs)
        return CtlHyperlinkFixed(ctl=ctrl)

    # endregion Static Methods

    # region Properties
    @property
    def model(self) -> UnoControlFixedHyperlinkModel:
        # pylint: disable=no-member
        return cast("UnoControlFixedHyperlinkModel", super().model)

    @property
    def model_ex(self) -> ModelHyperlinkFixed:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_hyperlink_fixed import ModelHyperlinkFixed

            self._model_ex = ModelHyperlinkFixed(self.model)
        return self._model_ex

    @property
    def view(self) -> UnoControlFixedHyperlink:
        # pylint: disable=no-member
        return cast("UnoControlFixedHyperlink", super().view)

    @property
    def view_ex(self) -> ViewFixedHyperlink:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        # pylint: disable=no-member
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_fixed_hyperlink import ViewFixedHyperlink

            self._view_ex = ViewFixedHyperlink(self.view)
        return self._view_ex

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_hyperlink_fixed import ModelHyperlinkFixed
    from ooodev.dialog.dl_control.view.view_fixed_hyperlink import ViewFixedHyperlink
