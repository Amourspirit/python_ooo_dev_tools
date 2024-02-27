from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from com.sun.star.awt import XControl
from com.sun.star.container import XIndexContainer
from com.sun.star.form import XGridColumnFactory

from ooodev.adapter.container.container_events import ContainerEvents
from ooodev.adapter.form.grid_control_events import GridControlEvents
from ooodev.adapter.form.reset_events import ResetEvents
from ooodev.adapter.form.update_events import UpdateEvents
from ooodev.adapter.script.script_events import ScriptEvents
from ooodev.adapter.view.selection_change_events import SelectionChangeEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.form_component_kind import FormComponentKind
from ooodev.loader import lo as mLo

from ooodev.form.controls.form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.form.component import GridControl as ControlModel  # service
    from com.sun.star.form.control import GridControl as ControlView  # service
    from ooodev.units.unit_obj import UnitT
    from ooodev.loader.inst.lo_inst import LoInst


class FormCtlGrid(
    FormCtlBase, ContainerEvents, GridControlEvents, ResetEvents, UpdateEvents, ScriptEvents, SelectionChangeEvents
):
    """``com.sun.star.form.component.GridControl`` control"""

    def __init__(self, ctl: XControl, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            ctl (XControl): Control supporting ``com.sun.star.form.component.GridControl`` service.
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

    # region Grid
    def get_column_types(self) -> tuple[str, ...]:
        """Gets the available column types"""
        return self.model.getColumnTypes()

    def create_grid_column(self, data_field: str, col_kind: str, width: int | UnitT = 0) -> None:
        """
        Adds a column to the gird

        Args:
            data_field (str): the database field to which the column should be bound
            col_kind (str):  the column type such as "NumericField"
            width (int, UnitT): the column width (in mm) or a ``UnitT`` implementation.
                If 0, no width is set. Defaults to ``0``.

        Returns:
            None:

        See Also:
            :py:meth:`~.form_ctl_grid.FormCtlGrid.get_column_types`
        """
        # column container and factory
        grid_model = self.model
        col_container = mLo.Lo.qi(XIndexContainer, grid_model, True)
        col_factory = mLo.Lo.qi(XGridColumnFactory, grid_model, True)

        # create the column
        col_props = col_factory.createColumn(col_kind)
        col_props.setPropertyValue("DataField", data_field)
        col_props.setPropertyValue("Label", data_field)
        col_props.setPropertyValue("Name", data_field)
        try:
            width_val = cast(int, width.get_value_mm100())  # type: ignore
        except AttributeError:
            width_val = cast(int, width) * 10  # type: ignore
        if width_val > 0:
            col_props.setPropertyValue("Width", width_val)

        # add properties column to container
        col_container.insertByIndex(col_container.getCount(), col_props)

    def get_current_column_position(self) -> int:
        """Gets the current column position"""
        return self.view.getCurrentColumnPosition()

    def set_current_column_position(self, position: int) -> None:
        """Sets the current column position"""
        self.view.setCurrentColumnPosition(position)

    # endregion Grid

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
