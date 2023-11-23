from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from com.sun.star.awt import XControl

from ooodev.adapter.form.submission.submission_veto_events import SubmissionVetoEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.form_component_kind import FormComponentKind

from .form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.form.component import SubmitButton as ControlModel  # service
    from com.sun.star.form.control import SubmitButton as ControlView  # service


class FormCtlSubmitButton(FormCtlBase, SubmissionVetoEvents):
    """``com.sun.star.form.component.SubmitButton`` control"""

    def __init__(self, ctl: XControl) -> None:
        FormCtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        SubmissionVetoEvents.__init__(self, trigger_args=generic_args, cb=self._on_submission_veto_add_remove)

    # region Lazy Listeners
    def _on_submission_veto_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addSubmissionVetoListener(self.events_listener_submission_veto)
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
        return FormComponentKind.SUBMIT_BUTTON

    # endregion Overrides

    # region Properties

    @property
    def model(self) -> ControlModel:
        """Gets the model for this control"""
        return self.get_model()

    @property
    def step(self) -> int:
        """Gets/Sets the step"""
        return self.model.Step

    @step.setter
    def step(self, value: int) -> None:
        self.model.Step = value

    @property
    def view(self) -> ControlView:
        """Gets the view of this control"""
        return self.get_view()

    # endregion Properties
