from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.document.document_event_events import DocumentEventEvents
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.adapter.view.print_job_events import PrintJobEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs


if TYPE_CHECKING:
    from com.sun.star.sheet import SpreadsheetDocument  # service


class SpreadsheetDocumentComp(
    ComponentBase, DocumentEventEvents, ModifyEvents, PrintJobEvents, PropertyChangeImplement, VetoableChangeImplement
):
    """
    Class for managing Spreadsheet Document Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: SpreadsheetDocument) -> None:
        """
        Constructor

        Args:
            component (SpreadsheetDocument): UNO Spreadsheet Document Component
        """
        ComponentBase.__init__(self, component)
        generic_args = self._get_generic_args()
        DocumentEventEvents.__init__(self, trigger_args=generic_args, cb=self._on_document_event_add_remove)
        ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        PrintJobEvents.__init__(self, trigger_args=generic_args, cb=self._on_print_job_add_remove)
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Lazy Listeners
    def _on_document_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addDocumentEventListener(self.events_listener_document_event)
        event.remove_callback = True

    def _on_modify_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addModifyListener(self.events_listener_modify)
        event.remove_callback = True

    def _on_print_job_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addPrintJobListener(self.events_listener_print_job)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SpreadsheetDocument",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> SpreadsheetDocument:
        """Spreadsheet Document Component"""
        return cast("SpreadsheetDocument", self._get_component())

    # endregion Properties
