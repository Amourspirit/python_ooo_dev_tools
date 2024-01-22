from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, List, overload, Iterable
import uno

from ooodev.adapter.document.document_event_events import DocumentEventEvents
from ooodev.adapter.presentation.presentation_document_comp import PresentationDocumentComp
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.adapter.view.print_job_events import PrintJobEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import draw as mDraw
from ooodev.utils import lo as mLo
from ooodev.utils.inst.lo.doc_type import DocType
from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils.partial.gui_partial import GuiPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.type_var import PathOrStr
from .partial.draw_doc_partial import DrawDocPartial
from . import impress_page as mImpressPage
from . import master_draw_page as mMasterDrawPage
from .draw_pages import DrawPages

if TYPE_CHECKING:
    from com.sun.star.beans import PropertyValue
    from com.sun.star.drawing import XDrawPage
    from com.sun.star.drawing import XDrawPages
    from com.sun.star.frame import XComponentLoader
    from com.sun.star.lang import XComponent
    from com.sun.star.presentation import XPresentation2
    from com.sun.star.presentation import XSlideShowController


class ImpressDoc(
    DrawDocPartial["ImpressDoc"],
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
    def __init__(self, doc: XComponent, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            self._lo_inst = mLo.Lo.current_lo
        else:
            self._lo_inst = lo_inst
        DrawDocPartial.__init__(self, owner=self, component=doc, lo_inst=self._lo_inst)
        PresentationDocumentComp.__init__(self, doc)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        DocumentEventEvents.__init__(self, trigger_args=generic_args, cb=self._on_document_event_add_remove)
        ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        PrintJobEvents.__init__(self, trigger_args=generic_args, cb=self._on_print_job_add_remove)
        QiPartial.__init__(self, component=doc, lo_inst=self._lo_inst)
        PropPartial.__init__(self, component=doc, lo_inst=self._lo_inst)
        GuiPartial.__init__(self, component=doc, lo_inst=self._lo_inst)
        ServicePartial.__init__(self, component=doc, lo_inst=self._lo_inst)
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
        return mImpressPage.ImpressPage(owner=self, component=result, lo_inst=self._lo_inst)

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
        return mImpressPage.ImpressPage(owner=self, component=page, lo_inst=self._lo_inst)

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
            return mImpressPage.ImpressPage(owner=self, component=result, lo_inst=self._lo_inst)
        if "slides" not in kwargs:
            kwargs["doc"] = self.component
        result = mDraw.Draw.get_slide(**kwargs)
        return mImpressPage.ImpressPage(owner=self, component=result, lo_inst=self._lo_inst)

    # endregion get_slide()

    def get_slides_list(self) -> List[mImpressPage.ImpressPage[ImpressDoc]]:
        """
        Gets all the slides as a list of XDrawPage

        Returns:
            List[ImpressPage[_T]]: List of pages
        """
        slides = mDraw.Draw.get_slides_list(self.component)
        return [mImpressPage.ImpressPage(owner=self, component=slide, lo_inst=self._lo_inst) for slide in slides]

    def get_viewed_page(self) -> mImpressPage.ImpressPage[ImpressDoc]:
        """
        Gets viewed page

        Raises:
            DrawPageError: If error occurs.

        Returns:
            ImpressPage: Draw Page
        """
        page = mDraw.Draw.get_viewed_page(self.component)
        return mImpressPage.ImpressPage(owner=self, component=page, lo_inst=self._lo_inst)

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
        return mMasterDrawPage.MasterDrawPage(owner=self, component=page, lo_inst=self._lo_inst)

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
        return mImpressPage.ImpressPage(owner=self, component=page, lo_inst=self._lo_inst)

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
        return mDraw.Draw.get_show_controller(show)

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
        return mImpressPage.ImpressPage(owner=self, component=slide, lo_inst=self._lo_inst)

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

        .. versionadded:: 0.20.2
        """
        return self._lo_inst.save_doc(self.component, fnm, password, format)  # type: ignore

    # endregion save_doc

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
            Exception: if unable to open document
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
        doc = lo_inst.open_doc(fnm=fnm, loader=loader, props=props)  # type: ignore
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
            lo_inst (LoInst): Lo instance.

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
            lo_inst (LoInst): Lo instance.

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
            lo_inst (LoInst, Optional): Lo instance.

        Raises:
            Exception: if unable to open document.

        Returns:
            ImpressDoc: Document.

        See Also:
            - :ref:`ch02_open_doc`
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        if loader is None:
            doc = lo_inst.open_readonly_doc(fnm=fnm)
        else:
            doc = lo_inst.open_readonly_doc(fnm=fnm, loader=loader)
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
            lo_inst (LoInst, Optional): Lo instance.

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
            lo_inst (LoInst, Optional): Lo instance.

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
            lo_inst (LoInst, Optional): Lo instance.

        Returns:
            ImpressDoc: Document

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
        return ImpressDoc(doc=doc, lo_inst=lo_inst)

    # endregion open_flat_doc()
    # endregion Static Open Methods

    # region Properties
    @property
    def slides(self) -> DrawPages[ImpressDoc]:
        """
        Returns:
            Any: Draw Pages.
        """
        if self._pages is None:
            self._pages = DrawPages(owner=self, slides=self.component.getDrawPages(), lo_inst=self._lo_inst)
        return cast("DrawPages[ImpressDoc]", self._pages)

    # endregion Properties
