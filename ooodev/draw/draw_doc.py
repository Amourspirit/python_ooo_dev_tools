from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, overload
import uno
from com.sun.star.util import XCloseable

from ooodev.adapter.document.document_event_events import DocumentEventEvents
from ooodev.adapter.drawing.drawing_document_comp import DrawingDocumentComp
from ooodev.adapter.frame.storable2_partial import Storable2Partial
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.adapter.view.print_job_events import PrintJobEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils import lo as mLo
from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils.partial.gui_partial import GuiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.adapter.util.close_events import CloseEvents
from ooodev.utils.type_var import PathOrStr
from .draw_pages import DrawPages
from .partial.draw_doc_partial import DrawDocPartial

if TYPE_CHECKING:
    from com.sun.star.lang import XComponent
    from com.sun.star.frame import XComponentLoader


class DrawDoc(
    DrawDocPartial["DrawDoc"],
    DrawingDocumentComp,
    DocumentEventEvents,
    ModifyEvents,
    PrintJobEvents,
    CloseEvents,
    Storable2Partial,
    QiPartial,
    PropPartial,
    GuiPartial,
    StylePartial,
):
    def __init__(self, doc: XComponent, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            self._lo_inst = mLo.Lo.current_lo
        else:
            self._lo_inst = lo_inst

        DrawDocPartial.__init__(self, owner=self, component=doc, lo_inst=self._lo_inst)
        DrawingDocumentComp.__init__(self, doc)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        DocumentEventEvents.__init__(self, trigger_args=generic_args, cb=self._on_document_event_add_remove)
        # ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        PrintJobEvents.__init__(self, trigger_args=generic_args, cb=self._on_print_job_add_remove)
        CloseEvents.__init__(self, trigger_args=generic_args, cb=self._on_print_job_add_remove)
        Storable2Partial.__init__(self, component=doc, interface=None)  # type: ignore
        QiPartial.__init__(self, component=doc, lo_inst=self._lo_inst)
        PropPartial.__init__(self, component=doc, lo_inst=self._lo_inst)
        GuiPartial.__init__(self, component=doc, lo_inst=self._lo_inst)
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

    def _on_close_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addCloseListener(self.events_listener_close)  # type: ignore
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
        return self._lo_inst.save_doc(self.component, fnm, password, format)  # type: ignore

    # endregion save_doc

    def close(self, deliver_ownership=True) -> None:
        """
        Try to close the Document.

        Nobody can guarantee real closing of called object - because it can disagree with that if any still running processes can't be canceled yet.
        It's not allowed to block this call till internal operations will be finished here.

        See Also:
            See LibreOffice API: `XCloseable.close() <https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1util_1_1XCloseable.html#af3d34677f1707b1904f8e07be4408592>`__
        """
        self.qi(XCloseable, True).close(deliver_ownership)

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

    @property
    def lo_inst(self) -> LoInst:
        """
        Returns:
            LoInst: LibreOffice instance.
        """
        return self._lo_inst

    # endregion Properties
