from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from com.sun.star.awt import XControl

from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.adapter.form.approve_action_events import ApproveActionEvents
from ooodev.adapter.form.reset_events import ResetEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.form_component_kind import FormComponentKind

from .form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.form.component import CommandButton as ControlModel  # service
    from com.sun.star.form.control import CommandButton as ControlView  # service


class FormCtlButton(FormCtlBase, ActionEvents, ApproveActionEvents, ResetEvents):
    """``com.sun.star.form.component.CommandButton`` control"""

    def __init__(self, ctl: XControl) -> None:
        FormCtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        ActionEvents.__init__(self, trigger_args=generic_args, cb=self._on_action_events_listener_add_remove)
        ApproveActionEvents.__init__(self, trigger_args=generic_args, cb=self._on_approve_action_add_remove)
        ResetEvents.__init__(self, trigger_args=generic_args, cb=self._on_reset_add_remove)

    # region Lazy Listeners
    def _on_action_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addActionListener(self.events_listener_action)
        event.remove_callback = True

    def _on_approve_action_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addApproveActionListener(self.events_listener_approve_action)
        event.remove_callback = True

    def _on_reset_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.model.addResetListener(self.events_listener_reset)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides

    if TYPE_CHECKING:
        # override the methods to provide type hinting
        def get_view(self) -> ControlView:
            """Gets the view of this control"""
            return cast("ControlView", super().get_view())

        def get_model(self) -> ControlModel:
            """Gets the model for this control"""
            return cast("ControlModel", super().get_model())

    def get_form_component_kind(self) -> FormComponentKind:
        """Gets the kind of form component this control is"""
        return FormComponentKind.COMMAND_BUTTON

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> ControlView:
        """Gets the view of this control"""
        return self.get_view()

    @property
    def model(self) -> ControlModel:
        """Gets the model for this control"""
        return self.get_model()

    # endregion Properties
