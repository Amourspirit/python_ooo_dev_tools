from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, overload, Iterable
import uno
from com.sun.star.frame import XComponentLoader
from com.sun.star.util import XCloseable

from ooodev.adapter.document.document_event_events import DocumentEventEvents
from ooodev.adapter.drawing.drawing_document_comp import DrawingDocumentComp
from ooodev.adapter.frame.storable2_partial import Storable2Partial
from ooodev.adapter.util.close_events import CloseEvents
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.adapter.view.print_job_events import PrintJobEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_partial import StylePartial
from ooodev.utils import info as mInfo
from ooodev.utils import lo as mLo
from ooodev.utils.context.lo_context import LoContext
from ooodev.utils.inst.lo.doc_type import DocType
from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils.inst.lo.service import Service as LoService
from ooodev.utils.partial.gui_partial import GuiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.type_var import PathOrStr
from .draw_pages import DrawPages
from .partial.draw_doc_partial import DrawDocPartial

if TYPE_CHECKING:
    from com.sun.star.lang import XComponent
    from com.sun.star.beans import PropertyValue


class DrawDoc(
    DrawDocPartial["DrawDoc"],
    LoInstPropsPartial,
    DrawingDocumentComp,
    DocumentEventEvents,
    ModifyEvents,
    PrintJobEvents,
    CloseEvents,
    Storable2Partial,
    QiPartial,
    PropPartial,
    GuiPartial,
    ServicePartial,
    StylePartial,
):
    """Draw document Class"""

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
        """
        return self.lo_inst.save_doc(self.component, fnm, password, format)  # type: ignore

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

    # region Create Document
    @overload
    @staticmethod
    def create_doc() -> DrawDoc:
        """
        Creates a new Draw document.

        Returns:
            DrawDoc: DrawDoc representing document
        """
        ...

    @overload
    @staticmethod
    def create_doc(loader: XComponentLoader) -> DrawDoc:
        """
        Creates a new Draw document.

        Args:
            loader (XComponentLoader): Component Loader. Usually generated with :py:class:`~.lo.Lo`

        Returns:
            DrawDoc: DrawDoc representing document
        """
        ...

    @overload
    @staticmethod
    def create_doc(lo_inst: LoInst) -> DrawDoc:
        """
        Creates a new Draw document.

        Args:
            lo_inst (LoInst): Lo instance.

        Returns:
            DrawDoc: DrawDoc representing document
        """
        ...

    @staticmethod
    def create_doc(*args, **kwargs) -> DrawDoc:
        """
        Creates a new Draw document.

        Args:
            loader (XComponentLoader, optional): Component Loader. Usually generated with :py:class:`~.lo.Lo`
            lo_inst (LoInst, optional): Lo instance.

        Returns:
            DrawDoc: DrawDoc representing document
        """
        doc = None
        lo_inst = None
        # 0 or 1 args
        arguments = list(args)
        arguments.extend(kwargs.values())
        count = len(arguments)
        if count == 0:
            doc = mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.DRAW)
        if count == 1:
            arg = arguments[0]
            if mLo.Lo.is_uno_interfaces(arg, XComponentLoader):
                doc = mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.DRAW, loader=arg)
            elif isinstance(arg, LoInst):
                lo_inst = arg
                with LoContext(arg):
                    doc = lo_inst.create_doc(doc_type=mLo.Lo.DocTypeStr.DRAW)
        if doc is None:
            raise TypeError("create_doc() got an unexpected argument")
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        return DrawDoc(doc=doc, lo_inst=lo_inst)

    # endregion Create Document

    # region create_doc_from_template()

    @overload
    @staticmethod
    def create_doc_from_template(template_path: PathOrStr) -> DrawDoc:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.

        Returns:
            DrawDoc: Document as DrawDoc instance.
        """
        ...

    @overload
    @staticmethod
    def create_doc_from_template(template_path: PathOrStr, *, lo_inst: LoInst) -> DrawDoc:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            DrawDoc: Document as DrawDoc instance.
        """
        ...

    @overload
    @staticmethod
    def create_doc_from_template(template_path: PathOrStr, loader: XComponentLoader) -> DrawDoc:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.
            loader (XComponentLoader): Component Loader.

        Returns:
            DrawDoc: Document as DrawDoc instance.
        """
        ...

    @overload
    @staticmethod
    def create_doc_from_template(template_path: PathOrStr, loader: XComponentLoader, lo_inst: LoInst) -> DrawDoc:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            DrawDoc: Document as DrawDoc instance.
        """
        ...

    @staticmethod
    def create_doc_from_template(
        template_path: PathOrStr, loader: XComponentLoader | None = None, lo_inst: LoInst | None = None
    ) -> DrawDoc:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.
            loader (XComponentLoader, optional): Component Loader.
            lo_inst (LoInst, optional): Lo instance. Used when created multiple documents.

        Raises:
            Exception: If unable to create document.

        Returns:
            DrawDoc: Document as DrawDoc instance.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        if loader is None:
            doc = lo_inst.create_doc_from_template(template_path=template_path)
        else:
            doc = lo_inst.create_doc_from_template(template_path=template_path, loader=loader)
        return DrawDoc(doc=doc, lo_inst=lo_inst)

    # endregion create_doc_from_template()

    # region Static Open Methods
    # region open_doc()
    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr) -> DrawDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open

        Returns:
            DrawDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, *, lo_inst: LoInst | None) -> DrawDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            lo_inst (LoInst): Lo instance.

        Returns:
            DrawDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, loader: XComponentLoader) -> DrawDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader

        Returns:
            DrawDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, loader: XComponentLoader, *, lo_inst: LoInst) -> DrawDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader
            lo_inst (LoInst): Lo instance.

        Returns:
            DrawDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, *, props: Iterable[PropertyValue]) -> DrawDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            props (Iterable[PropertyValue]): Properties passed to component loader

        Returns:
            DrawDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, *, props: Iterable[PropertyValue], lo_inst: LoInst) -> DrawDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            props (Iterable[PropertyValue]): Properties passed to component loader
            lo_inst (LoInst): Lo instance.

        Returns:
            DrawDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, loader: XComponentLoader, props: Iterable[PropertyValue]) -> DrawDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader
            props (Iterable[PropertyValue]): Properties passed to component loader
            lo_inst (LoInst): Lo instance.


        Returns:
            DrawDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, loader: XComponentLoader, props: Iterable[PropertyValue], lo_inst: LoInst) -> DrawDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader
            props (Iterable[PropertyValue]): Properties passed to component loader
            lo_inst (LoInst): Lo instance.

        Returns:
            DrawDoc: Document
        """
        ...

    @staticmethod
    def open_doc(
        fnm: PathOrStr,
        loader: XComponentLoader | None = None,
        props: Iterable[PropertyValue] | None = None,
        lo_inst: LoInst | None = None,
    ) -> DrawDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader, optional): Component Loader
            props (Iterable[PropertyValue], optional): Properties passed to component loader
            lo_inst (LoInst, Optional): Lo instance.

        Raises:
            CancelEventError: if DOC_OPENING event is canceled.

        Returns:
            DrawDoc: Document

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_OPENING` :eventref:`src-docs-event-cancel`
                - :py:attr:`~.events.lo_named_event.LoNamedEvent.DOC_OPENED` :eventref:`src-docs-event`

        Note:
            Event args ``event_data`` is a dictionary containing all method parameters.

        See Also:
            - :py:meth:`~Lo.open_doc`
            - :py:meth:`load_office`
            - :ref:`ch02_open_doc`

        Note:
            If connection it office is a remote server then File URL must be used,
            such as ``file:///home/user/fancy.odt``
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        doc = lo_inst.open_doc(fnm=fnm, loader=loader, props=props)  # type: ignore
        return DrawDoc(doc=doc, lo_inst=lo_inst)

    # endregion open_doc()

    # region open_readonly_doc()
    @overload
    @staticmethod
    def open_readonly_doc(fnm: PathOrStr) -> DrawDoc:
        """
        Open a office document as read-only

        Args:
            fnm (PathOrStr): path of document to open.

        Returns:
            DrawDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_readonly_doc(fnm: PathOrStr, *, lo_inst: LoInst) -> DrawDoc:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            DrawDoc: Document.
        """
        ...

    @overload
    @staticmethod
    def open_readonly_doc(fnm: PathOrStr, loader: XComponentLoader) -> DrawDoc:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.

        Returns:
            DrawDoc: Document.
        """
        ...

    @overload
    @staticmethod
    def open_readonly_doc(fnm: PathOrStr, loader: XComponentLoader, lo_inst: LoInst) -> DrawDoc:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            DrawDoc: Document.
        """
        ...

    @staticmethod
    def open_readonly_doc(
        fnm: PathOrStr, loader: XComponentLoader | None = None, lo_inst: LoInst | None = None
    ) -> DrawDoc:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst, optional): Lo instance. Used when created multiple documents.

        Raises:
            Exception: if unable to open document.

        Returns:
            DrawDoc: Document.

        See Also:
            - :ref:`ch02_open_doc`
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        if loader is None:
            doc = lo_inst.open_readonly_doc(fnm=fnm)
        else:
            doc = lo_inst.open_readonly_doc(fnm=fnm, loader=loader)
        return DrawDoc(doc=doc, lo_inst=lo_inst)

    # endregion open_readonly_doc()

    # region open_flat_doc()
    @overload
    @staticmethod
    def open_flat_doc(fnm: PathOrStr, doc_type: DocType) -> DrawDoc:
        """
        Opens a flat document

        Args:
            fnm (PathOrStr): path of XML document
            doc_type (DocType): Type of document to open

        Returns:
            DrawDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_flat_doc(fnm: PathOrStr, doc_type: DocType, *, lo_inst: LoInst) -> DrawDoc:
        """
        Opens a flat document

        Args:
            fnm (PathOrStr): path of XML document
            doc_type (DocType): Type of document to open
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            DrawDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_flat_doc(fnm: PathOrStr, doc_type: DocType, loader: XComponentLoader) -> DrawDoc:
        """
        Opens a flat document

        Args:
            fnm (PathOrStr): path of XML document
            doc_type (DocType): Type of document to open
            loader (XComponentLoader, optional): Component loader

        Returns:
            DrawDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_flat_doc(fnm: PathOrStr, doc_type: DocType, loader: XComponentLoader, lo_inst: LoInst) -> DrawDoc:
        """
        Opens a flat document

        Args:
            fnm (PathOrStr): path of XML document
            doc_type (DocType): Type of document to open
            loader (XComponentLoader, optional): Component loader
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            DrawDoc: Document
        """
        ...

    @staticmethod
    def open_flat_doc(
        fnm: PathOrStr, doc_type: DocType, loader: XComponentLoader | None = None, lo_inst: LoInst | None = None
    ) -> DrawDoc:
        """
        Opens a flat document

        Args:
            fnm (PathOrStr): path of XML document
            doc_type (DocType): Type of document to open
            loader (XComponentLoader, optional): Component loader
            lo_inst (LoInst, optional): Lo instance. Used when created multiple documents.

        Raises:
            Exception: if unable to open document.

        Returns:
            DrawDoc: Document

        See Also:
            - :py:meth:`~Lo.open_flat_doc`
            - :ref:`ch02_open_doc`
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        if loader is None:
            doc = lo_inst.open_flat_doc(fnm=fnm, doc_type=doc_type)
        else:
            doc = lo_inst.open_flat_doc(fnm=fnm, doc_type=doc_type, loader=loader)
        return DrawDoc(doc=doc, lo_inst=lo_inst)

    # endregion open_flat_doc()
    # endregion Static Open Methods

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
