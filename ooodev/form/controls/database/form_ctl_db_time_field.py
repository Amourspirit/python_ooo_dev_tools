from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from com.sun.star.awt import XControl

from ooodev.utils.kind.form_component_kind import FormComponentKind
from ooodev.adapter.form.update_events import UpdateEvents

from ..form_ctl_time_field import FormCtlTimeField

if TYPE_CHECKING:
    from com.sun.star.form.component import DatabaseTimeField as ControlModel  # service
    from ooodev.events.args.listener_event_args import ListenerEventArgs


class FormCtlDbTimeField(FormCtlTimeField, UpdateEvents):
    """``com.sun.star.form.component.DatabaseTimeField`` control"""

    def __init__(self, ctl: XControl) -> None:
        FormCtlTimeField.__init__(self, ctl)
        generic_args = self._get_generic_args()
        UpdateEvents.__init__(self, trigger_args=generic_args, cb=self._on_update_events_add_remove)

    # region Lazy Listeners
    def _on_update_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.get_model().addUpdateListener(self.events_listener_update)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides

    if TYPE_CHECKING:
        # override the methods to provide type hinting
        def get_model(self) -> ControlModel:
            """Gets the model for this control"""
            return cast("ControlModel", super().get_model())

    def get_form_component_kind(self) -> FormComponentKind:
        """Gets the kind of form component this control is"""
        return FormComponentKind.DATABASE_TIME_FIELD

    # endregion Overrides

    # region Properties
    if TYPE_CHECKING:
        # override the properties to provide type hinting
        @property
        def model(self) -> ControlModel:
            """Gets the model for this control"""
            return self.get_model()

    # endregion Properties
