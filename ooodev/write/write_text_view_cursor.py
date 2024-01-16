from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, TypeVar, Generic
import uno

from com.sun.star.text import XTextViewCursor

if TYPE_CHECKING:
    from com.sun.star.text import XTextDocument

from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_view_cursor_comp import TextViewCursorComp
from ooodev.adapter.view.line_cursor_partial import LineCursorPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import write as mWrite
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils import selection as mSelection
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.type_var import PathOrStr
from ooodev.write import write_doc as mWriteDoc
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.write import WriteNamedEvent

from .partial.text_cursor_partial import TextCursorPartial

T = TypeVar("T", bound="ComponentT")


class WriteTextViewCursor(
    Generic[T],
    TextCursorPartial,
    TextViewCursorComp,
    LineCursorPartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
    PropPartial,
    QiPartial,
    StylePartial,
    EventsPartial,
):
    """Represents a writer text view cursor."""

    def __init__(self, owner: T, component: XTextViewCursor) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextViewCursor): A UNO object that supports ``com.sun.star.text.TextViewCursor`` service.
        """
        self.__owner = owner
        TextCursorPartial.__init__(self, owner=owner, component=component)
        TextViewCursorComp.__init__(self, component)  # type: ignore
        LineCursorPartial.__init__(self, component, None)  # type: ignore
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        PropPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)
        QiPartial.__init__(self, component=component, lo_inst=mLo.Lo.current_lo)  # type: ignore
        StylePartial.__init__(self, component=component)
        EventsPartial.__init__(self)

    def __len__(self) -> int:
        return mSelection.Selection.range_len(cast("XTextDocument", self.owner.component), self.component)

    def get_coord_str(self) -> str:
        """
        Gets coordinates for cursor in format such as ``"10, 10"``

        Returns:
            str: coordinates as string
        """
        return mWrite.Write.get_coord_str(self.component)  # type: ignore

    def get_current_page_num(self) -> int:
        """
        Gets current page number.

        Returns:
            int: current page number
        """
        return mWrite.Write.get_current_page(self.component)  # type: ignore

    get_current_page = get_current_page_num

    def get_text_view_cursor(self) -> XTextViewCursor:
        """
        Gets the text view cursor.

        Returns:
            XTextViewCursor: text view cursor
        """
        return self.qi(XTextViewCursor, True)

    # region export images
    def export_page_png(self, fnm: PathOrStr = "", resolution: int = 96) -> None:
        """
        Exports doc pages as png images.

        Args:
            fnm (PathOrStr, optional): Image file name.
            resolution (int, optional): Resolution in dpi. Defaults to 96.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.EXPORTING_PAGE_PNG` :eventref:`src-docs-event-cancel-export`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.EXPORTED_PAGE_PNG` :eventref:`src-docs-event-export`

        Returns:
            None:

        Note:
            On exporting event is :ref:`cancel_event_args_export`.
            On exported event is :ref:`event_args_export`.
            Args ``event_data`` is a :py:class:`~ooodev.write.filter.export_png.ExportPngT` dictionary.

            If ``fnm`` is not specified, the image file name is created based on the document name and page number
            and written to the same folder as the document.

        See Also:
            :py:class:`~ooodev.write.export.page_png.PagePng`
        """
        if not isinstance(self.owner, mWriteDoc.WriteDoc):
            raise TypeError(f"Owner must be of type {mWriteDoc.WriteDoc.__name__}")

        def on_exporting(source: Any, args: Any) -> None:
            self.trigger_event(WriteNamedEvent.EXPORTING_PAGE_PNG, args)

        def on_exported(source: Any, args: Any) -> None:
            self.trigger_event(WriteNamedEvent.EXPORTED_PAGE_PNG, args)

        # don't import to avoid unnecessary dependencies and overhead until needed.
        from .export.page_png import PagePng

        doc = self.owner
        exporter = PagePng(doc)
        exporter.subscribe_event_exporting(on_exporting)
        exporter.subscribe_event_exported(on_exported)

        exporter.export(fnm=fnm, resolution=resolution)

    def export_page_jpg(self, fnm: PathOrStr = "", resolution: int = 96) -> None:
        """
        Exports doc pages as jpg images.

        Args:
            fnm (PathOrStr, optional): Image file name.
            resolution (int, optional): Resolution in dpi. Defaults to 96.

        :events:
            .. cssclass:: lo_event

                - :py:attr:`~.events.write_named_event.WriteNamedEvent.EXPORTING_PAGE_JPG` :eventref:`src-docs-event-cancel-export`
                - :py:attr:`~.events.write_named_event.WriteNamedEvent.EXPORTED_PAGE_JPG` :eventref:`src-docs-event-export`

        Returns:
            None:

        Note:
            On exporting event is :ref:`cancel_event_args_export`.
            On exported event is :ref:`event_args_export`.
            Args ``event_data`` is a :py:class:`~ooodev.write.filter.export_jpg.ExportJpgT` dictionary.

            If ``fnm`` is not specified, the image file name is created based on the document name and page number
            and written to the same folder as the document.

        See Also:
            :py:class:`~ooodev.write.export.page_jpg.PageJpg`
        """
        if not isinstance(self.owner, mWriteDoc.WriteDoc):
            raise TypeError(f"Owner must be of type {mWriteDoc.WriteDoc.__name__}")

        def on_exporting(source: Any, args: Any) -> None:
            self.trigger_event(WriteNamedEvent.EXPORTING_PAGE_JPG, args)

        def on_exported(source: Any, args: Any) -> None:
            self.trigger_event(WriteNamedEvent.EXPORTED_PAGE_JPG, args)

        # don't import to avoid unnecessary dependencies and overhead until needed.
        from .export.page_jpg import PageJpg

        doc = self.owner
        exporter = PageJpg(doc)
        exporter.subscribe_event_exporting(on_exporting)
        exporter.subscribe_event_exported(on_exported)

        exporter.export(fnm=fnm, resolution=resolution)

    # endregion export images

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self.__owner

    # endregion Properties
