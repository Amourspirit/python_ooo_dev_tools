from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, overload
import uno

from ooodev.adapter.document.document_event_events import DocumentEventEvents
from ooodev.adapter.drawing.drawing_document_comp import DrawingDocumentComp
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.adapter.view.print_job_events import PrintJobEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils import lo as mLo
from ooodev.utils.type_var import PathOrStr
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from .draw_pages import DrawPages
from .partial.draw_doc_partial import DrawDocPartial

if TYPE_CHECKING:
    from com.sun.star.lang import XComponent


class DrawDoc(
    DrawDocPartial["DrawDoc"],
    DrawingDocumentComp,
    DocumentEventEvents,
    ModifyEvents,
    PrintJobEvents,
    QiPartial,
    PropPartial,
    StylePartial,
):
    def __init__(self, doc: XComponent) -> None:
        DrawDocPartial.__init__(self, owner=self, component=doc)
        DrawingDocumentComp.__init__(self, doc)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        DocumentEventEvents.__init__(self, trigger_args=generic_args, cb=self._on_document_event_add_remove)
        # ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        PrintJobEvents.__init__(self, trigger_args=generic_args, cb=self._on_print_job_add_remove)
        QiPartial.__init__(self, component=doc, lo_inst=mLo.Lo.current_lo)
        PropPartial.__init__(self, component=doc, lo_inst=mLo.Lo.current_lo)
        StylePartial.__init__(self, component=doc)
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

    # endregion Lazy Listeners

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

    # region save_doc

    @overload
    def save_doc(self, fnm: PathOrStr) -> bool:
        """
        Save document.

        Args:
            fnm (PathOrStr): file path to save as.

        Returns:
            bool: ``False`` if ``DOC_SAVING`` event is canceled; Otherwise, ``True``
        """
        ...

    @overload
    def save_doc(self, fnm: PathOrStr, password: str) -> bool:
        """
        Save document.

        Args:
            fnm (PathOrStr): file path to save as.
            password (str): Password to save document with.


        Returns:
            bool: ``False`` if ``DOC_SAVING`` event is canceled; Otherwise, ``True``
        """
        ...

    @overload
    def save_doc(self, fnm: PathOrStr, password: str, format: str) -> bool:  # pylint: disable=W0622
        """
        Save document.

        Args:
            fnm (PathOrStr): file path to save as.
            password (str): Password to save document with.
            format (str): document format such as 'odt' or 'xml'.

        Returns:
            bool: ``False`` if ``DOC_SAVING`` event is canceled; Otherwise, ``True``.
        """
        ...

    def save_doc(self, fnm: PathOrStr, password: str | None = None, format: str | None = None) -> bool:
        """
        Save document.

        Args:
            fnm (PathOrStr): file path to save as.
            password (str, optional): password to save document with.
            format (str, optional): document format such as 'odt' or 'xml'.

        Raises:
            MissingInterfaceError: If doc does not implement XStorable interface.

        Returns:
            bool: ``False`` if DOC_SAVING event is canceled; Otherwise, ``True``

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_SAVING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_SAVED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing all method parameters.

        Attention:
            :py:meth:`~.utils.lo.Lo.store_doc` method is called along with any of its events.

        See Also:
            :ref:`ch02_save_doc`

        .. versionadded:: 0.20.2
        """
        return mLo.Lo.save_doc(self.component, fnm, password, format)  # type: ignore

    # endregion save_doc

    # region Properties

    @property
    def slides(self) -> DrawPages[DrawDoc]:
        """
        Returns:
            Any: Draw Pages.
        """
        if self._pages is None:
            self._pages = DrawPages(owner=self, slides=self.component.getDrawPages())
        return cast("DrawPages[DrawDoc]", self._pages)

    # endregion Properties
