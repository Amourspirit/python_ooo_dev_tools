from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from com.sun.star.awt import XControl

from ooodev.adapter.form.submission.submission_veto_events import SubmissionVetoEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.form_component_kind import FormComponentKind

from ooodev.form.controls.form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.form.component import SubmitButton as ControlModel  # service
    from com.sun.star.form.control import SubmitButton as ControlView  # service
    from ooodev.loader.inst.lo_inst import LoInst


class FormCtlSubmitButton(FormCtlBase, SubmissionVetoEvents):
    """``com.sun.star.form.component.SubmitButton`` control"""

    def __init__(self, ctl: XControl, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            ctl (XControl): Control supporting ``com.sun.star.form.component.SubmitButton`` service.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to ``None``.

        Returns:
            None:

        Note:
            If the :ref:`LoContext <ooodev.utils.context.lo_context.LoContext>` manager is use before this class is instantiated,
            then the Lo instance will be set using the current Lo instance. That the context manager has set.
            Generally speaking this means that there is no need to set ``lo_inst`` when instantiating this class.

        See Also:
            :ref:`ooodev.form.Forms`.
        """
        FormCtlBase.__init__(self, ctl=ctl, lo_inst=lo_inst)
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
