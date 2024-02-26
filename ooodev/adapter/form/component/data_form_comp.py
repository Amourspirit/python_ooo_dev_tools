from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.form.component.form_comp import FormComp
from ooodev.adapter.form.reset_partial import ResetPartial
from ooodev.adapter.form.loadable_partial import LoadablePartial
from ooodev.adapter.form.load_events import LoadEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.sdbc.result_set_update_partial import ResultSetUpdatePartial
from ooodev.adapter.sdbcx.delete_rows_partial import DeleteRowsPartial
from ooodev.adapter.sdb.result_set_access_partial import ResultSetAccessPartial
from ooodev.adapter.sdb.parameters_supplier_partial import ParametersSupplierPartial

if TYPE_CHECKING:
    from com.sun.star.form.component import DataForm  # service


class DataFormComp(
    FormComp,
    ResetPartial,
    LoadablePartial,
    ResultSetUpdatePartial,
    DeleteRowsPartial,
    ResultSetAccessPartial,
    ParametersSupplierPartial,
    LoadEvents,
):
    """
    Class for managing DataForm Service.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: Any) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Component that implements ``com.sun.star.form.component.DataForm`` service.
        """

        FormComp.__init__(self, component)
        ResetPartial.__init__(self, component=self.component, interface=None)
        LoadablePartial.__init__(self, component=self.component, interface=None)
        ResultSetUpdatePartial.__init__(self, component=self.component, interface=None)
        DeleteRowsPartial.__init__(self, component=self.component, interface=None)
        ResultSetAccessPartial.__init__(self, component=self.component, interface=None)
        ParametersSupplierPartial.__init__(self, component=self.component, interface=None)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        LoadEvents.__init__(self, trigger_args=generic_args, cb=self._on_event_load_add_remove)

    # region Lazy Listeners

    def _on_event_load_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addLoadListener(self.events_listener_load)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.form.component.DataForm",)

    # endregion Overrides

    # region Properties

    if TYPE_CHECKING:

        @property
        def component(self) -> DataForm:
            """DataForm Component"""
            # pylint: disable=no-member
            return cast("DataForm", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
