from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, List, overload, Iterable
import uno
from com.sun.star.frame import XComponentLoader

from ooodev.adapter.document.document_event_events import DocumentEventEvents
from ooodev.adapter.presentation.presentation_document_comp import PresentationDocumentComp
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.adapter.view.print_job_events import PrintJobEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import draw as mDraw
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
from .partial.draw_doc_partial import DrawDocPartial
from . import impress_page as mImpressPage
from . import master_draw_page as mMasterDrawPage
from .impress_pages import ImpressPages

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.drawing import XDrawPage
    from com.sun.star.drawing import XDrawPages
    from com.sun.star.lang import XComponent
    from com.sun.star.presentation import XPresentation2
    from com.sun.star.presentation import XSlideShowController


class ImpressDoc(
    DrawDocPartial["ImpressDoc"],
    LoInstPropsPartial,
    PresentationDocumentComp,
    DocumentEventEvents,
    ModifyEvents,
    PrintJobEvents,
    QiPartial,
    PropPartial,
    GuiPartial,
    ServicePartial,
    StylePartial,
):
    """Impress Document Class"""

    def __init__(self, doc: XComponent, lo_inst: LoInst | None = None) -> None:
        """
        Constructor.

        Args:
            doc (XComponent): Impress Document component.
            lo_inst (LoInst, optional): Lo instance. Used when created multiple documents. Defaults to None.

        Raises:
            NotSupportedDocumentError: If not an Impress Document.
        """
        if not mInfo.Info.is_doc_type(doc, LoService.IMPRESS):
            raise mEx.NotSupportedDocumentError("Document is not a Impress document")
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        DrawDocPartial.__init__(self, owner=self, component=doc, lo_inst=self.lo_inst)
        PresentationDocumentComp.__init__(self, doc)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        DocumentEventEvents.__init__(self, trigger_args=generic_args, cb=self._on_document_event_add_remove)
        ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        PrintJobEvents.__init__(self, trigger_args=generic_args, cb=self._on_print_job_add_remove)
        QiPartial.__init__(self, component=doc, lo_inst=self.lo_inst)
        PropPartial.__init__(self, component=doc, lo_inst=self.lo_inst)
        GuiPartial.__init__(self, component=doc, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=doc, lo_inst=self.lo_inst)
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

    # region DrawDocPartial Overrides

    def get_slides(self) -> ImpressPages[ImpressDoc]:
        """
        Gets the impress pages of a document.

        Returns:
            ImpressPages[ImpressDoc]: Impress Pages.
        """
        return self.slides

    # endregion DrawDocPartial Overrides

    def add_slide(self) -> mImpressPage.ImpressPage[ImpressDoc]:
        """
        Add a slide to the end of the document.

        Raises:
            DrawPageMissingError: If unable to get pages.
            DrawPageError: If any other error occurs.

        Returns:
            mImpressPage: The slide that was inserted at the end of the document.
        """
        result = mDraw.Draw.add_slide(doc=self.component)
        return mImpressPage.ImpressPage(owner=self, component=result, lo_inst=self.lo_inst)

    def duplicate(self, idx: int) -> mImpressPage.ImpressPage[ImpressDoc]:
        """
        Duplicates a slide

        Args:
            idx (int): Index of slide to duplicate.

        Raises:
            DrawError If unable to create duplicate.

        Returns:
            ImpressPage: Duplicated slide.
        """
        page = mDraw.Draw.duplicate(self.component, idx)
        return mImpressPage.ImpressPage(owner=self, component=page, lo_inst=self.lo_inst)

    # region get_slide()
    @overload
    def get_slide(self) -> mImpressPage.ImpressPage[ImpressDoc]:
        """
        Gets draw page by page at index ``0``.

        Returns:
            ImpressPage: Draw Page.
        """
        ...

    @overload
    def get_slide(self, *, idx: int) -> mImpressPage.ImpressPage[ImpressDoc]:
        """
        Gets draw page by page index

        Args:
            idx (int): Index of draw page. Default ``0``

        Returns:
            ImpressPage: Draw Page.
        """
        ...

    @overload
    def get_slide(self, *, slides: XDrawPages) -> mImpressPage.ImpressPage[ImpressDoc]:
        """
        Gets draw page at index ``0`` from ``slides``.

        Args:
            slides (XDrawPages): Draw Pages

        Returns:
            ImpressPage: Draw Page.
        """
        ...

    @overload
    def get_slide(self, *, slides: XDrawPages, idx: int) -> mImpressPage.ImpressPage[ImpressDoc]:
        """
        Gets draw page by page index from ``slides``.

        Args:
            slides (XDrawPages): Draw Pages
            idx (int): Index of slide. Default ``0``

        Returns:
            ImpressPage: Slide as Draw Page.
        """
        ...

    def get_slide(self, **kwargs) -> mImpressPage.ImpressPage[ImpressDoc]:
        """
        Gets slide

        Args:
            slides (XDrawPages): Draw Pages
            idx (int): Index of slide. Default ``0``

        Raises:
            IndexError: If ``idx`` is out of bounds
            DrawError: If any other error occurs.

        Returns:
            ImpressPage: Slide as Draw Page.
        """
        if not kwargs:
            result = mDraw.Draw.get_slide(doc=self.component)
            return mImpressPage.ImpressPage(owner=self, component=result, lo_inst=self.lo_inst)
        if "slides" not in kwargs:
            kwargs["doc"] = self.component
        result = mDraw.Draw.get_slide(**kwargs)
        return mImpressPage.ImpressPage(owner=self, component=result, lo_inst=self.lo_inst)

    # endregion get_slide()

    def get_slides_list(self) -> List[mImpressPage.ImpressPage[ImpressDoc]]:
        """
        Gets all the slides as a list of XDrawPage

        Returns:
            List[ImpressPage[_T]]: List of pages
        """
        slides = mDraw.Draw.get_slides_list(self.component)
        return [mImpressPage.ImpressPage(owner=self, component=slide, lo_inst=self.lo_inst) for slide in slides]

    def get_viewed_page(self) -> mImpressPage.ImpressPage[ImpressDoc]:
        """
        Gets viewed page

        Raises:
            DrawPageError: If error occurs.

        Returns:
            ImpressPage: Draw Page
        """
        page = mDraw.Draw.get_viewed_page(self.component)
        return mImpressPage.ImpressPage(owner=self, component=page, lo_inst=self.lo_inst)

    def get_handout_master_page(self) -> mMasterDrawPage.MasterDrawPage[ImpressDoc]:
        """
        Gets handout master page for an impress document.

        Raises:
            DrawError: If unable to get hand-out master page.
            DrawPageMissingError: If Draw Page is ``None``.

        Returns:
            MasterDrawPage: Impress Page
        """
        page = mDraw.Draw.get_handout_master_page(self.component)
        return mMasterDrawPage.MasterDrawPage(owner=self, component=page, lo_inst=self.lo_inst)

    def get_notes_page_by_index(self, idx: int) -> mImpressPage.ImpressPage[ImpressDoc]:
        """
        Gets notes page by index.

        Each draw page has a notes page.

        Args:
            idx (int): Index

        Raises:
            DrawPageError: If error occurs.

        Returns:
            ImpressPage: Notes Page.

        See Also:
            :py:meth:`~.draw.Draw.get_notes_page`
        """
        page = mDraw.Draw.get_notes_page_by_index(self.component, idx)
        return mImpressPage.ImpressPage(owner=self, component=page, lo_inst=self.lo_inst)

    def get_show(self) -> XPresentation2:
        """
        Gets Slide show Presentation.

        Raises:
            DrawError: If error occurs.

        Returns:
            XPresentation2: Slide Show Presentation.
        """
        return mDraw.Draw.get_show(self.component)

    def get_show_controller(self) -> XSlideShowController:
        """
        Gets slide show controller

        Args:
            show (XPresentation2): Slide Show Presentation

        Raises:
            DrawError: If error occurs.

        Returns:
            XSlideShowController: Slide Show Controller.

        Note:
            It may take a little bit for the slides show to start.
            For this reason this method will wait up to five seconds.
        """
        show = self.get_show()
        result = mDraw.Draw.get_show_controller(show)
        return result

    def insert_slide(self, idx: int) -> mImpressPage.ImpressPage[ImpressDoc]:
        """
        Inserts a slide at the given position in the document

        Args:
            idx (int): Index

        Raises:
            DrawPageMissingError: If unable to get pages.
            DrawPageError: If any other error occurs.

        Returns:
            DrawPage: New slide that was inserted.
        """
        slide = mDraw.Draw.insert_slide(doc=self.component, idx=idx)
        return mImpressPage.ImpressPage(owner=self, component=slide, lo_inst=self.lo_inst)

    def remove_master_page(self, slide: XDrawPage) -> None:
        """
        Removes a master page.

        Args:
            slide (XDrawPage): Draw page to remove.

        Raises:
            DrawError: If unable to remove master page.

        Returns:
            None:
        """
        mDraw.Draw.remove_master_page(doc=self.component, slide=slide)

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

    # region Create Document
    @overload
    @staticmethod
    def create_doc() -> ImpressDoc:
        """
        Creates a new Impress document.

        Returns:
            ImpressDoc: ImpressDoc representing document
        """
        ...

    @overload
    @staticmethod
    def create_doc(loader: XComponentLoader) -> ImpressDoc:
        """
        Creates a new Impress document.

        Args:
            loader (XComponentLoader): Component Loader. Usually generated with :py:class:`~.lo.Lo`

        Returns:
            ImpressDoc: ImpressDoc representing document
        """
        ...

    @overload
    @staticmethod
    def create_doc(lo_inst: LoInst) -> ImpressDoc:
        """
        Creates a new Impress document.

        Args:
            lo_inst (LoInst): Lo instance.

        Returns:
            ImpressDoc: ImpressDoc representing document
        """
        ...

    @staticmethod
    def create_doc(*args, **kwargs) -> ImpressDoc:
        """
        Creates a new Impress document.

        Args:
            loader (XComponentLoader, optional): Component Loader. Usually generated with :py:class:`~.lo.Lo`
            lo_inst (LoInst, optional): Lo instance.

        Returns:
            ImpressDoc: ImpressDoc representing document
        """
        doc = None
        lo_inst = None
        # 0 or 1 args
        arguments = list(args)
        arguments.extend(kwargs.values())
        count = len(arguments)
        if count == 0:
            doc = mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.IMPRESS)
        if count == 1:
            arg = arguments[0]
            if mLo.Lo.is_uno_interfaces(arg, XComponentLoader):
                doc = mLo.Lo.create_doc(doc_type=mLo.Lo.DocTypeStr.IMPRESS, loader=arg)
            if isinstance(arg, LoInst):
                with LoContext(arg) as lo_inst:
                    doc = lo_inst.create_doc(doc_type=mLo.Lo.DocTypeStr.IMPRESS)
        if doc is None:
            raise TypeError("create_doc() got an unexpected argument")
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        return ImpressDoc(doc=doc, lo_inst=lo_inst)

    # endregion Create Document

    # region create_doc_from_template()

    @overload
    @staticmethod
    def create_doc_from_template(template_path: PathOrStr) -> ImpressDoc:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.

        Returns:
            ImpressDoc: Document as ImpressDoc instance.
        """
        ...

    @overload
    @staticmethod
    def create_doc_from_template(template_path: PathOrStr, *, lo_inst: LoInst) -> ImpressDoc:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.
            lo_inst (LoInst, optional): Lo instance. Used when created multiple documents.

        Returns:
            ImpressDoc: Document as ImpressDoc instance.
        """
        ...

    @overload
    @staticmethod
    def create_doc_from_template(template_path: PathOrStr, loader: XComponentLoader) -> ImpressDoc:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.
            loader (XComponentLoader, optional): Component Loader.

        Returns:
            ImpressDoc: Document as ImpressDoc instance.
        """
        ...

    @overload
    @staticmethod
    def create_doc_from_template(template_path: PathOrStr, loader: XComponentLoader, lo_inst: LoInst) -> ImpressDoc:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.
            loader (XComponentLoader, optional): Component Loader.
            lo_inst (LoInst, optional): Lo instance. Used when created multiple documents.

        Returns:
            ImpressDoc: Document as ImpressDoc instance.
        """
        ...

    @staticmethod
    def create_doc_from_template(
        template_path: PathOrStr, loader: XComponentLoader | None = None, lo_inst: LoInst | None = None
    ) -> ImpressDoc:
        """
        Create a document from a template.

        Args:
            template_path (PathOrStr): path to template file.
            loader (XComponentLoader, optional): Component Loader.
            lo_inst (LoInst, optional): Lo instance. Used when created multiple documents.

        Raises:
            Exception: If unable to create document.

        Returns:
            ImpressDoc: Document as ImpressDoc instance.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        with LoContext(lo_inst) as inst:
            if loader is None:
                doc = inst.create_doc_from_template(template_path=template_path)
            else:
                doc = inst.create_doc_from_template(template_path=template_path, loader=loader)
        return ImpressDoc(doc=doc, lo_inst=lo_inst)

    # endregion create_doc_from_template()

    # region Static Open Methods
    # region open_doc()
    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr) -> ImpressDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open

        Returns:
            ImpressDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, *, lo_inst: LoInst | None) -> ImpressDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            lo_inst (LoInst): Lo instance.

        Returns:
            ImpressDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, loader: XComponentLoader) -> ImpressDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader

        Returns:
            ImpressDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, loader: XComponentLoader, *, lo_inst: LoInst) -> ImpressDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader
            lo_inst (LoInst): Lo instance.

        Returns:
            ImpressDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, *, props: Iterable[PropertyValue]) -> ImpressDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            props (Iterable[PropertyValue]): Properties passed to component loader

        Returns:
            ImpressDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, *, props: Iterable[PropertyValue], lo_inst: LoInst) -> ImpressDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            props (Iterable[PropertyValue]): Properties passed to component loader
            lo_inst (LoInst): Lo instance.

        Returns:
            ImpressDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(fnm: PathOrStr, loader: XComponentLoader, props: Iterable[PropertyValue]) -> ImpressDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader
            props (Iterable[PropertyValue]): Properties passed to component loader
            lo_inst (LoInst): Lo instance.


        Returns:
            ImpressDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_doc(
        fnm: PathOrStr, loader: XComponentLoader, props: Iterable[PropertyValue], lo_inst: LoInst
    ) -> ImpressDoc:
        """
        Open a office document

        Args:
            fnm (PathOrStr): path of document to open
            loader (XComponentLoader): Component Loader
            props (Iterable[PropertyValue]): Properties passed to component loader
            lo_inst (LoInst): Lo instance.

        Returns:
            ImpressDoc: Document
        """
        ...

    @staticmethod
    def open_doc(
        fnm: PathOrStr,
        loader: XComponentLoader | None = None,
        props: Iterable[PropertyValue] | None = None,
        lo_inst: LoInst | None = None,
    ) -> ImpressDoc:
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
            ImpressDoc: Document

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
        with LoContext(lo_inst) as inst:
            doc = inst.open_doc(fnm=fnm, loader=loader, props=props)  # type: ignore

        return ImpressDoc(doc=doc, lo_inst=lo_inst)

    # endregion open_doc()

    # region open_readonly_doc()
    @overload
    @staticmethod
    def open_readonly_doc(fnm: PathOrStr) -> ImpressDoc:
        """
        Open a office document as read-only

        Args:
            fnm (PathOrStr): path of document to open.

        Returns:
            ImpressDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_readonly_doc(fnm: PathOrStr, *, lo_inst: LoInst) -> ImpressDoc:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            ImpressDoc: Document.
        """
        ...

    @overload
    @staticmethod
    def open_readonly_doc(fnm: PathOrStr, loader: XComponentLoader) -> ImpressDoc:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.

        Returns:
            ImpressDoc: Document.
        """
        ...

    @overload
    @staticmethod
    def open_readonly_doc(fnm: PathOrStr, loader: XComponentLoader, lo_inst: LoInst) -> ImpressDoc:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            ImpressDoc: Document.
        """
        ...

    @staticmethod
    def open_readonly_doc(
        fnm: PathOrStr, loader: XComponentLoader | None = None, lo_inst: LoInst | None = None
    ) -> ImpressDoc:
        """
        Open a office document as read-only.

        Args:
            fnm (PathOrStr): path of document to open.
            loader (XComponentLoader): Component Loader.
            lo_inst (LoInst, optional): Lo instance. Used when created multiple documents.

        Returns:
            ImpressDoc: Document.

        See Also:
            - :ref:`ch02_open_doc`
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        with LoContext(lo_inst) as inst:
            if loader is None:
                doc = inst.open_readonly_doc(fnm=fnm)
            else:
                doc = inst.open_readonly_doc(fnm=fnm, loader=loader)
        return ImpressDoc(doc=doc, lo_inst=lo_inst)

    # endregion open_readonly_doc()

    # region open_flat_doc()
    @overload
    @staticmethod
    def open_flat_doc(fnm: PathOrStr, doc_type: DocType) -> ImpressDoc:
        """
        Opens a flat document

        Args:
            fnm (PathOrStr): path of XML document
            doc_type (DocType): Type of document to open

        Returns:
            ImpressDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_flat_doc(fnm: PathOrStr, doc_type: DocType, *, lo_inst: LoInst) -> ImpressDoc:
        """
        Opens a flat document

        Args:
            fnm (PathOrStr): path of XML document
            doc_type (DocType): Type of document to open
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            ImpressDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_flat_doc(fnm: PathOrStr, doc_type: DocType, loader: XComponentLoader) -> ImpressDoc:
        """
        Opens a flat document

        Args:
            fnm (PathOrStr): path of XML document
            doc_type (DocType): Type of document to open
            loader (XComponentLoader, optional): Component loader

        Returns:
            ImpressDoc: Document
        """
        ...

    @overload
    @staticmethod
    def open_flat_doc(fnm: PathOrStr, doc_type: DocType, loader: XComponentLoader, lo_inst: LoInst) -> ImpressDoc:
        """
        Opens a flat document

        Args:
            fnm (PathOrStr): path of XML document
            doc_type (DocType): Type of document to open
            loader (XComponentLoader, optional): Component loader
            lo_inst (LoInst): Lo instance. Used when created multiple documents.

        Returns:
            ImpressDoc: Document
        """
        ...

    @staticmethod
    def open_flat_doc(
        fnm: PathOrStr, doc_type: DocType, loader: XComponentLoader | None = None, lo_inst: LoInst | None = None
    ) -> ImpressDoc:
        """
        Opens a flat document

        Args:
            fnm (PathOrStr): path of XML document
            doc_type (DocType): Type of document to open
            loader (XComponentLoader, optional): Component loader
            lo_inst (LoInst, optional): Lo instance. Used when created multiple documents.

        Returns:
            ImpressDoc: Document

        See Also:
            - :py:meth:`~Lo.open_flat_doc`
            - :ref:`ch02_open_doc`
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        with LoContext(lo_inst) as inst:
            if loader is None:
                doc = inst.open_flat_doc(fnm=fnm, doc_type=doc_type)
            else:
                doc = inst.open_flat_doc(fnm=fnm, doc_type=doc_type, loader=loader)
        return ImpressDoc(doc=doc, lo_inst=lo_inst)

    # endregion open_flat_doc()
    # endregion Static Open Methods

    # region Properties
    @property
    def slides(self) -> ImpressPages[ImpressDoc]:
        """
        Returns:
            Any: Draw Pages.
        """
        if self._pages is None:
            self._pages = ImpressPages(owner=self, slides=self.component.getDrawPages(), lo_inst=self.lo_inst)
        return cast("ImpressPages[ImpressDoc]", self._pages)

    # endregion Properties
