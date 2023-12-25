from __future__ import annotations
from typing import Any, TYPE_CHECKING, List, overload
import uno

from ooodev.adapter.document.document_event_events import DocumentEventEvents
from ooodev.adapter.presentation.presentation_document_comp import PresentationDocumentComp
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.adapter.view.print_job_events import PrintJobEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import draw as mDraw
from ooodev.utils import lo as mLo
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from .partial.draw_doc_partial import DrawDocPartial
from . import impress_page as mImpressPage
from . import master_draw_page as mMasterDrawPage

if TYPE_CHECKING:
    from com.sun.star.lang import XComponent
    from com.sun.star.drawing import XDrawPage
    from com.sun.star.drawing import XDrawPages
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
    StylePartial,
):
    def __init__(self, doc: XComponent) -> None:
        DrawDocPartial.__init__(self, owner=self, component=doc)
        PresentationDocumentComp.__init__(self, doc)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        DocumentEventEvents.__init__(self, trigger_args=generic_args, cb=self._on_document_event_add_remove)
        ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        PrintJobEvents.__init__(self, trigger_args=generic_args, cb=self._on_print_job_add_remove)
        QiPartial.__init__(self, component=doc, lo_inst=mLo.Lo.current_lo)
        PropPartial.__init__(self, component=doc, lo_inst=mLo.Lo.current_lo)
        StylePartial.__init__(self, component=doc)

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
        return mImpressPage.ImpressPage(self, result)

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
        return mImpressPage.ImpressPage(self, page)

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
            return mImpressPage.ImpressPage(self, result)
        if not "slides" in kwargs:
            kwargs["doc"] = self.component
        result = mDraw.Draw.get_slide(**kwargs)
        return mImpressPage.ImpressPage(self, result)

    # endregion get_slide()

    def get_slides_list(self) -> List[mImpressPage.ImpressPage[ImpressDoc]]:
        """
        Gets all the slides as a list of XDrawPage

        Returns:
            List[ImpressPage[_T]]: List of pages
        """
        slides = mDraw.Draw.get_slides_list(self.component)
        return [mImpressPage.ImpressPage(self, slide) for slide in slides]

    def get_viewed_page(self) -> mImpressPage.ImpressPage[ImpressDoc]:
        """
        Gets viewed page

        Raises:
            DrawPageError: If error occurs.

        Returns:
            ImpressPage: Draw Page
        """
        page = mDraw.Draw.get_viewed_page(self.component)
        return mImpressPage.ImpressPage(self, page)

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
        return mMasterDrawPage.MasterDrawPage(self, page)

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
        return mImpressPage.ImpressPage(self, page)

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
        return mImpressPage.ImpressPage(self, slide)

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
