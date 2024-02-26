from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.adapter.document.document_event_events import DocumentEventEvents
from ooodev.adapter.drawing.drawing_document_comp import DrawingDocumentComp
from ooodev.adapter.frame.storable2_partial import Storable2Partial
from ooodev.adapter.util.close_events import CloseEvents
from ooodev.adapter.view.print_job_events import PrintJobEvents
from ooodev.dialog.partial.create_dialog_partial import CreateDialogPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.doc_type import DocType
from ooodev.loader.inst.service import Service as LoService
from ooodev.utils import info as mInfo
from ooodev.utils.partial.dispatch_partial import DispatchPartial
from ooodev.utils.partial.doc_io_partial import DocIoPartial
from ooodev.utils.partial.gui_partial import GuiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.draw.draw_pages import DrawPages
from ooodev.draw.partial.draw_doc_partial import DrawDocPartial

if TYPE_CHECKING:
    from com.sun.star.lang import XComponent
    from ooodev.loader.inst.lo_inst import LoInst

# pylint: disable=unused-argument


class DrawDoc(
    DrawDocPartial["DrawDoc"],
    LoInstPropsPartial,
    DrawingDocumentComp,
    DocumentEventEvents,
    PrintJobEvents,
    CloseEvents,
    Storable2Partial,
    QiPartial,
    PropPartial,
    GuiPartial,
    ServicePartial,
    StylePartial,
    EventsPartial,
    DocIoPartial["DrawDoc"],
    CreateDialogPartial,
    DispatchPartial,
):
    """Draw document Class"""

    DOC_TYPE: DocType = DocType.DRAW

    def __init__(self, doc: XComponent, lo_inst: LoInst | None = None) -> None:
        """
        Constructor.

        Args:
            doc (XComponent): Writer Document component.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.

        Raises:
            NotSupportedDocumentError: If not a valid Draw document.

        Returns:
            None:
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo

        if not mInfo.Info.is_doc_type(doc, LoService.DRAW):
            raise mEx.NotSupportedDocumentError("Document is not a Draw document")

        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        DrawDocPartial.__init__(self, owner=self, component=doc, lo_inst=self.lo_inst)
        DrawingDocumentComp.__init__(self, doc)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        DocumentEventEvents.__init__(self, trigger_args=generic_args, cb=self._on_document_event_add_remove)
        # ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        PrintJobEvents.__init__(self, trigger_args=generic_args, cb=self._on_print_job_add_remove)
        CloseEvents.__init__(self, trigger_args=generic_args, cb=self._on_print_job_add_remove)
        Storable2Partial.__init__(self, component=doc, interface=None)  # type: ignore
        QiPartial.__init__(self, component=doc, lo_inst=self.lo_inst)
        PropPartial.__init__(self, component=doc, lo_inst=self.lo_inst)
        GuiPartial.__init__(self, component=doc, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=doc)
        ServicePartial.__init__(self, component=doc, lo_inst=self.lo_inst)
        EventsPartial.__init__(self)
        DocIoPartial.__init__(self, owner=self, lo_inst=self.lo_inst)
        CreateDialogPartial.__init__(self, lo_inst=self.lo_inst)
        DispatchPartial.__init__(self, lo_inst=self.lo_inst, events=self)
        self._pages = None

    # region Lazy Listeners

    def _on_modify_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addModifyListener(self.events_listener_modify)
        event.remove_callback = True

    def _on_document_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addDocumentEventListener(self.events_listener_document_event)
        event.remove_callback = True

    def _on_print_job_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addPrintJobListener(self.events_listener_print_job)
        event.remove_callback = True

    def _on_close_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addCloseListener(self.events_listener_close)  # type: ignore
        event.remove_callback = True

    # endregion Lazy Listeners

    # region context manage
    def __enter__(self) -> DrawDoc:
        self.lock_controllers()
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.unlock_controllers()

    # endregion context manage

    # region Overrides
    def get_slides(self) -> DrawPages:
        """
        Gets the draw pages of a document.

        Args:
            doc (XComponent): Document.

        Raises:
            DrawPageMissingError: If there are no draw pages.
            DrawPageError: If any other error occurs.

        Returns:
            DrawPages: Draw Pages.
        """
        return self.slides

    def delete_slide(self, idx: int) -> bool:
        """
        Deletes a slide

        Args:
            idx (int): Index. Can be a negative value to delete from the end of the document.
                For example, -1 will delete the last slide.

        Returns:
            bool: ``True`` on success; Otherwise, ``False``
        """
        if idx < 0:
            idx = len(self.slides) + idx
        if idx < 0:
            raise IndexError("list index out of range")

        return super().delete_slide(idx=idx)

    # endregion Overrides

    # region DocIoPartial Overrides
    # region from_current_doc()
    @classmethod
    def _on_from_current_doc_loading(cls, event_args: CancelEventArgs) -> None:
        """
        Event called while from_current_doc loading.

        Args:
            event_args (EventArgs): Event data.

        Returns:
            None:

        Note:
            event_args.event_data is a dictionary and contains the document in a key named 'doc'.
        """
        event_args.event_data["doc_type"] = cls.DOC_TYPE

    @classmethod
    def _on_from_current_doc_loaded(cls, event_args: EventArgs) -> None:
        """
        Event called after from_current_doc is called.

        Args:
            event_args (EventArgs): Event data.

        Returns:
            None:

        Note:
            event_args.event_data is a dictionary and contains the document in a key named 'doc'.
        """
        doc = cast(DrawDoc, event_args.event_data["doc"])
        if doc.DOC_TYPE != cls.DOC_TYPE:
            raise mEx.NotSupportedDocumentError(f"Document '{type(doc).__name__}' is not a Draw document.")

    # endregion from_current_doc()
    # endregion DocIoPartial Overrides

    # region Properties
    @property
    def slides(self) -> DrawPages[DrawDoc]:
        """
        Returns:
            Any: Draw Pages.
        """
        if self._pages is None:
            self._pages = DrawPages(owner=self, slides=self.component.getDrawPages(), lo_inst=self.lo_inst)
        return cast("DrawPages[DrawDoc]", self._pages)

    # endregion Properties
