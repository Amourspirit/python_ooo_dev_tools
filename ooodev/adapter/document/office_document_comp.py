from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.document.document_event_events import DocumentEventEvents
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.adapter.view.print_job_events import PrintJobEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.frame.model_partial import ModelPartial


if TYPE_CHECKING:
    from com.sun.star.document import OfficeDocument  # service
    from com.sun.star.lang import XComponent


class OfficeDocumentComp(ComponentBase, ModelPartial, DocumentEventEvents, ModifyEvents, PrintJobEvents):
    """
    Class for managing Sheet Cell Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XComponent) -> None:
        """
        Constructor

        Args:
            component (XComponent): UNO Component that supports ``com.sun.star.document.OfficeDocument`` service.
        """
        ComponentBase.__init__(self, component)
        ModelPartial.__init__(self, component=component, interface=None)  # type: ignore
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        DocumentEventEvents.__init__(self, trigger_args=generic_args, cb=self._on_document_event_add_remove)
        ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        PrintJobEvents.__init__(self, trigger_args=generic_args, cb=self._on_print_job_add_remove)

    # region Lazy Listeners
    def _on_modify_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addModifyListener(self.events_listener_modify)
        event.remove_callback = True

    def _on_print_job_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addPrintJobListener(self.events_listener_print_job)
        event.remove_callback = True

    def _on_document_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addDocumentEventListener(self.events_listener_document_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.document.OfficeDocument",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> OfficeDocument:
        """OfficeDocument Component"""
        return cast("OfficeDocument", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
