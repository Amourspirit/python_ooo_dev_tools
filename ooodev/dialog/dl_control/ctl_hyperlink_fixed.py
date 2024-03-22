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
from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedHyperlink  # service
    from com.sun.star.awt import UnoControlFixedHyperlinkModel  # service
    from ooodev.dialog.dl_control.model.model_hyperlink_fixed import ModelHyperlinkFixed
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

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_hyperlink_fixed import ModelHyperlinkFixed
