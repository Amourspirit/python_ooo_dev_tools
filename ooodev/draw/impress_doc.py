from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, List, overload
import uno

from ooodev.adapter.document.document_event_events import DocumentEventEvents
from ooodev.adapter.presentation.presentation_document_comp import PresentationDocumentComp
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.adapter.view.print_job_events import PrintJobEvents
from ooodev.dialog.partial.create_dialog_partial import CreateDialogPartial
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst import DocType
from ooodev.loader.inst import Service as LoService
from ooodev.office import draw as mDraw
from ooodev.utils import info as mInfo
from ooodev.utils.partial.dispatch_partial import DispatchPartial
from ooodev.utils.partial.doc_io_partial import DocIoPartial
from ooodev.utils.partial.gui_partial import GuiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from .partial.draw_doc_partial import DrawDocPartial
from . import impress_page as mImpressPage
from . import master_draw_page as mMasterDrawPage
from .impress_pages import ImpressPages

if TYPE_CHECKING:
    from com.sun.star.drawing import XDrawPage
    from com.sun.star.drawing import XDrawPages
    from com.sun.star.lang import XComponent
    from com.sun.star.presentation import XPresentation2
    from com.sun.star.presentation import XSlideShowController
    from ooodev.loader.inst import LoInst


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
    EventsPartial,
    StylePartial,
    DocIoPartial["ImpressDoc"],
    CreateDialogPartial,
    DispatchPartial,
):
    """Impress Document Class"""

    DOC_TYPE: DocType = DocType.IMPRESS

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
        EventsPartial.__init__(self)
        StylePartial.__init__(self, component=doc)
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

    # endregion Lazy Listeners

    # region context manage
    def __enter__(self) -> ImpressDoc:
        self.lock_controllers()
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.unlock_controllers()

    # endregion context manage

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
        doc = cast(ImpressDoc, event_args.event_data["doc"])
        if doc.DOC_TYPE != cls.DOC_TYPE:
            raise mEx.NotSupportedDocumentError(f"Document '{type(doc).__name__}' is not an Impress document.")

    # endregion from_current_doc()
    # endregion DocIoPartial Overrides

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
