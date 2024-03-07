from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, TypeVar, Generic
import uno

from com.sun.star.text import XTextViewCursor

from ooodev.mock import mock_g
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.style.character_properties_partial import CharacterPropertiesPartial
from ooodev.adapter.style.paragraph_properties_partial import ParagraphPropertiesPartial
from ooodev.adapter.text.text_view_cursor_comp import TextViewCursorComp
from ooodev.adapter.view.line_cursor_partial import LineCursorPartial
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import write as mWrite
from ooodev.loader import lo as mLo
from ooodev.utils import selection as mSelection
from ooodev.utils.context.lo_context import LoContext
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.type_var import PathOrStr
from ooodev.write import write_doc as mWriteDoc
from ooodev.events.write_named_event import WriteNamedEvent
from ooodev.write.partial.text_cursor_partial import TextCursorPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.proto.component_proto import ComponentT

if TYPE_CHECKING:
    from com.sun.star.text import XTextDocument
    from ooodev.write.style.direct.character_styler import CharacterStyler

T = TypeVar("T", bound="ComponentT")


class WriteTextViewCursor(
    LoInstPropsPartial,
    WriteDocPropPartial,
    TextCursorPartial[T],
    Generic[T],
    TextViewCursorComp,
    CharacterPropertiesPartial,
    ParagraphPropertiesPartial,
    LineCursorPartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
    PropPartial,
    QiPartial,
    StylePartial,
    EventsPartial,
):
    """Represents a writer text view cursor."""

    def __init__(self, owner: T, component: XTextViewCursor, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextViewCursor): A UNO object that supports ``com.sun.star.text.TextViewCursor`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        TextCursorPartial.__init__(self, owner=owner, component=component, lo_inst=self.lo_inst)
        TextViewCursorComp.__init__(self, component)  # type: ignore
        CharacterPropertiesPartial.__init__(self, component=component)  # type: ignore
        ParagraphPropertiesPartial.__init__(self, component=component)  # type: ignore
        LineCursorPartial.__init__(self, component, None)  # type: ignore
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)
        EventsPartial.__init__(self)
        self._style_direct_char = None

    def __len__(self) -> int:
        with LoContext(self.lo_inst):
            result = mSelection.Selection.range_len(cast("XTextDocument", self.owner.component), self.component)
        return result

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

        Example:

            .. code-block:: python

                from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
                from ooodev.events.args.event_args_export import EventArgsExport
                from ooodev.write import Write, WriteDoc
                from ooodev.write import WriteNamedEvent
                from ooodev.write.filter.export_png import ExportPngT

                # ... code omitted ...

                def export_pages_png(doc_file: str):
                    doc = WriteDoc(Write.open_doc(fnm=doc_file))

                    def on_exporting(source: Any, args: CancelEventArgsExport[ExportPngT]) -> None:
                        args.event_data["compression"] = 5 # default is 6

                    def on_exported(source: Any, args: EventArgsExport[ExportPngT]) -> None:
                        print(args.get("url"))

                    try:
                        view = doc.get_view_cursor()

                        doc_path = Path(doc.get_doc_path())
                        view.jump_to_last_page()
                        # optionally subscribe to events to make fine tune export
                        view.subscribe_event(WriteNamedEvent.EXPORTED_PAGE_PNG, on_exported)
                        view.subscribe_event(WriteNamedEvent.EXPORTING_PAGE_PNG, on_exporting)

                        for i in range(view.get_page(), 0, -1):
                            img_path = Path(tmp_path_fn, f"{doc_path.stem}_{i}.png")
                            # export 300 DPI
                            view.export_page_png(fnm=img_path, resolution=300)
                            view.jump_to_previous_page()

                        view.jump_to_first_page()
                    finally:
                        doc.close_doc()

        See Also:
            :py:class:`~ooodev.write.export.page_png.PagePng`
        """
        # pylint: disable=import-outside-toplevel
        # pylint: disable=unused-argument
        # pylint: disable=redefined-outer-name
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

        Example:

            .. code-block:: python

                from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
                from ooodev.events.args.event_args_export import EventArgsExport
                from ooodev.write import Write, WriteDoc
                from ooodev.write import WriteNamedEvent
                from ooodev.write.filter.export_jpg import ExportJpgT

                # ... code omitted ...

                def export_pages_jpg(doc_file: str):
                    doc = WriteDoc(Write.open_doc(fnm=doc_file))

                    def on_exporting(source: Any, args: CancelEventArgsExport[ExportJpgT]) -> None:
                        args.event_data["quality"] = 80 # Default is 75

                    def on_exported(source: Any, args: EventArgsExport[ExportJpgT]) -> None:
                        print(args.get("url"))

                    try:
                        view = doc.get_view_cursor()

                        doc_path = Path(doc.get_doc_path())
                        view.jump_to_last_page()
                        # optionally subscribe to events to make fine tune export
                        view.subscribe_event(WriteNamedEvent.EXPORTED_PAGE_JPG, on_exported)
                        view.subscribe_event(WriteNamedEvent.EXPORTING_PAGE_JPG, on_exporting)

                        for i in range(view.get_page(), 0, -1):
                            img_path = Path(tmp_path_fn, f"{doc_path.stem}_{i}.jpg")
                            # export 300 DPI
                            view.export_page_jpg(fnm=img_path, resolution=300)
                            view.jump_to_previous_page()

                        view.jump_to_first_page()
                    finally:
                        doc.close_doc()

        See Also:
            :py:class:`~ooodev.write.export.page_jpg.PageJpg`
        """
        # pylint: disable=import-outside-toplevel
        # pylint: disable=unused-argument
        # pylint: disable=redefined-outer-name
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
        return self._owner

    @property
    def style_direct_char(self) -> CharacterStyler:
        """
        Direct Character Styler.

        Returns:
            CharacterStyler: Character Styler
        """
        if self._style_direct_char is None:
            # pylint: disable=import-outside-toplevel
            from ooodev.write.style.direct.character_styler import CharacterStyler

            self._style_direct_char = CharacterStyler(write_doc=self.write_doc, component=self.component)
            self._style_direct_char.add_event_observers(self.event_observer)
        return self._style_direct_char

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.write.style.direct.character_styler import CharacterStyler
    from .export.page_png import PagePng
    from .export.page_jpg import PageJpg
