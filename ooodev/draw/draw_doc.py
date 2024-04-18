from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno

from ooodev.adapter.drawing.drawing_document_comp import DrawingDocumentComp
from ooodev.adapter.frame.storable2_partial import Storable2Partial
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.args.event_args import EventArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.loader.inst.clsid import CLSID
from ooodev.loader.inst.doc_type import DocType
from ooodev.loader.inst.service import Service as LoService
from ooodev.gui.menu.menu_app import MenuApp
from ooodev.gui.menu.menus import Menus
from ooodev.gui.menu.shortcuts import Shortcuts
from ooodev.utils import info as mInfo
from ooodev.draw.draw_pages import DrawPages
from ooodev.draw.partial.draw_doc_partial import DrawDocPartial
from ooodev.draw.draw_doc_view import DrawDocView
from ooodev.draw.partial.draw_doc_prop_partial import DrawDocPropPartial
from ooodev.draw.partial.doc_partial import DocPartial
from ooodev.office.partial.office_document_prop_partial import OfficeDocumentPropPartial

if TYPE_CHECKING:
    from com.sun.star.lang import XComponent
    from ooodev.loader.inst.lo_inst import LoInst

# pylint: disable=unused-argument


class DrawDoc(
    DrawingDocumentComp,
    OfficeDocumentPropPartial,
    DocPartial["DrawDoc"],
    DrawDocPropPartial,
    Storable2Partial,
):
    """Draw document Class"""

    DOC_TYPE: DocType = DocType.DRAW
    DOC_CLSID: CLSID = CLSID.DRAW

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
        # pylint: disable=non-parent-init-called
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo

        if not mInfo.Info.is_doc_type(doc, LoService.DRAW):
            raise mEx.NotSupportedDocumentError("Document is not a Draw document")
        DrawingDocumentComp.__init__(self, doc)
        OfficeDocumentPropPartial.__init__(self, office_doc=self)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        DocPartial.__init__(self, owner=self, component=doc, generic_args=generic_args, lo_inst=lo_inst)
        DrawDocPartial.__init__(self, owner=self, component=doc, lo_inst=self.lo_inst)
        DrawDocPropPartial.__init__(self, obj=self)
        # ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        Storable2Partial.__init__(self, component=doc, interface=None)  # type: ignore
        self._pages = None
        self._menu = None
        self._shortcuts = None

    # region Lazy Listeners

    def _on_modify_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addModifyListener(self.events_listener_modify)
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

    @property
    def current_controller(self) -> DrawDocView:
        """
        Gets controller from document.

        Returns:
            Any: controller.
        """
        comp = DrawDocView(owner=self, component=self.component.CurrentController)  # type: ignore
        return comp

    @property
    def menu(self) -> MenuApp:
        """
        Gets access to Draw Menus.

        Returns:
            MenuApp: Draw Menu

        Example:
            .. code-block:: python

                # Example of getting the Calc Menus
                file_menu = doc.menu["file"]
                file_menu[3].execute()

        .. versionadded:: 0.40.0
        """
        if self._menu is None:
            self._menu = Menus(lo_inst=self.lo_inst)[LoService.DRAW]
        return self._menu  # type: ignore

    @property
    def shortcuts(self) -> Shortcuts:
        """
        Gets access to Draw Shortcuts.

        Returns:
            Shortcuts: Draw Shortcuts

        .. versionadded:: 0.40.0
        """
        if self._shortcuts is None:
            self._shortcuts = Shortcuts(app=LoService.DRAW, lo_inst=self.lo_inst)
        return self._shortcuts

    # endregion Properties
