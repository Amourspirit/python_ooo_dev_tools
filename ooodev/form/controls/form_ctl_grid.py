from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from com.sun.star.awt import XControl

from ooodev.adapter.container.container_events import ContainerEvents
from ooodev.adapter.form.grid_control_events import GridControlEvents
from ooodev.adapter.form.reset_events import ResetEvents
from ooodev.adapter.form.update_events import UpdateEvents
from ooodev.adapter.script.script_events import ScriptEvents
from ooodev.adapter.view.selection_change_events import SelectionChangeEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.form_component_kind import FormComponentKind


from .form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.form.component import GridControl as ControlModel  # service
    from com.sun.star.form.control import GridControl as ControlView  # service


class FormCtlGrid(
    FormCtlBase, ContainerEvents, GridControlEvents, ResetEvents, UpdateEvents, ScriptEvents, SelectionChangeEvents
):
    """``com.sun.star.form.component.GridControl`` control"""

    def __init__(self, ctl: XControl) -> None:
        FormCtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        ResetEvents.__init__(self, trigger_args=generic_args, cb=self._on_reset_add_remove)
        GridControlEvents.__init__(self, trigger_args=generic_args, cb=self._on_grid_control_add_remove)
        SelectionChangeEvents.__init__(self, trigger_args=generic_args, cb=self._on_grid_selection_change_add_remove)
        UpdateEvents.__init__(self, trigger_args=generic_args, cb=self._on_update_add_remove)
        ContainerEvents.__init__(self, trigger_args=generic_args, cb=self._on_container_add_remove)
        ScriptEvents.__init__(self, trigger_args=generic_args, cb=self._on_script_add_remove)

    # region Lazy Listeners
    def _on_grid_control_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addGridControlListener(self.events_listener_grid_control)
        event.remove_callback = True

    def _on_update_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addUpdateListener(self.events_listener_update)
        event.remove_callback = True

    def _on_grid_selection_change_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addSelectionChangeListener(self.events_listener_selection_change)
        event.remove_callback = True

    def _on_reset_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.model.addResetListener(self.events_listener_reset)
        event.remove_callback = True

    def _on_container_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.model.addContainerListener(self.events_listener_container)
        event.remove_callback = True

    def _on_script_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.model.addScriptListener(self.events_listener_script)
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
    def enabled(self) -> bool:
        """Gets/Sets the enabled state for the control"""
        return self.model.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.model.Enabled = value

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
    def tab_stop(self) -> bool:
        """Gets/Sets the tab stop property"""
        return self.model.Tabstop

    @tab_stop.setter
    def tab_stop(self, value: bool) -> None:
        self.model.Tabstop = value

    @property
    def view(self) -> ControlView:
        """Gets the view of this control"""
        return self.get_view()

    # endregion Properties
