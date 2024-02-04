from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from com.sun.star.awt import XControl

from ooodev.adapter.form.data_aware_control_model_partial import DataAwareControlModelPartial
from ooodev.adapter.form.update_events import UpdateEvents
from ooodev.utils.kind.form_component_kind import FormComponentKind

from ..form_ctl_combo_box import FormCtlComboBox

if TYPE_CHECKING:
    from com.sun.star.form.component import DatabaseComboBox as ControlModel  # service
    from ooodev.events.args.listener_event_args import ListenerEventArgs
    from ooodev.loader.inst.lo_inst import LoInst


class FormCtlDbComboBox(FormCtlComboBox, DataAwareControlModelPartial, UpdateEvents):
    """``com.sun.star.form.component.DatabaseComboBox`` control"""

    def __init__(self, ctl: XControl, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            ctl (XControl): Control supporting ``com.sun.star.form.component.DatabaseComboBox`` service.
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
        FormCtlComboBox.__init__(self, ctl=ctl, lo_inst=lo_inst)
        generic_args = self._get_generic_args()
        UpdateEvents.__init__(self, trigger_args=generic_args, cb=self._on_update_events_add_remove)
        DataAwareControlModelPartial.__init__(self, self.get_model())

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
        return FormComponentKind.DATABASE_CURRENCY_FIELD

    # endregion Overrides

    # region Properties
    if TYPE_CHECKING:
        # override the properties to provide type hinting
        @property
        def model(self) -> ControlModel:
            """Gets the model for this control"""
            return self.get_model()

    # endregion Properties
