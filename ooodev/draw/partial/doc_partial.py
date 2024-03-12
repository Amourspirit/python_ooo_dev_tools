from __future__ import annotations
from typing import Any, cast, List, TYPE_CHECKING, TypeVar, Generic
import uno
from ooodev.adapter.document.document_event_events import DocumentEventEvents
from ooodev.adapter.container.index_access_comp import IndexAccessComp
from ooodev.adapter.util.close_events import CloseEvents
from ooodev.adapter.view.print_job_events import PrintJobEvents
from ooodev.dialog.partial.create_dialog_partial import CreateDialogPartial
from ooodev.draw.partial.draw_doc_partial import DrawDocPartial
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.utils.partial.dispatch_partial import DispatchPartial
from ooodev.utils.partial.doc_io_partial import DocIoPartial
from ooodev.utils.partial.gui_partial import GuiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils import info as mInfo


if TYPE_CHECKING:
    from com.sun.star.lang import XComponent
    from com.sun.star.drawing import GenericDrawingDocument
    from ooodev.adapter.drawing.generic_shape import GenericShapeComp
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.events.args.generic_args import GenericArgs
    from ooodev.proto.component_proto import ComponentT

_T = TypeVar("_T", bound="ComponentT")


class DocPartial(
    DrawDocPartial[_T],
    Generic[_T],
    LoInstPropsPartial,
    DocumentEventEvents,
    GuiPartial,
    PrintJobEvents,
    CloseEvents,
    ServicePartial,
    QiPartial,
    PropPartial,
    EventsPartial,
    DocIoPartial[_T],
    CreateDialogPartial,
    DispatchPartial,
    StylePartial,
):
    """
    Document partial class.


    This class represents common document operations for Impress and Draw documents.
    """

    def __init__(
        self, owner: _T, component: XComponent, generic_args: GenericArgs | None = None, lo_inst: LoInst | None = None
    ) -> None:
        self.__component = cast("GenericDrawingDocument", component)
        self.__owner = owner
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        GuiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        DrawDocPartial.__init__(self, owner=self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        DocumentEventEvents.__init__(self, trigger_args=generic_args, cb=self._on_document_event_add_remove)
        PrintJobEvents.__init__(self, trigger_args=generic_args, cb=self._on_print_job_add_remove)
        CloseEvents.__init__(self, trigger_args=generic_args, cb=self._on_close_add_remove)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        EventsPartial.__init__(self)
        DocIoPartial.__init__(self, owner=self, lo_inst=self.lo_inst)
        CreateDialogPartial.__init__(self, lo_inst=self.lo_inst)
        DispatchPartial.__init__(self, lo_inst=self.lo_inst, events=self)
        StylePartial.__init__(self, component=component)

    # region Lazy Listeners

    def _on_document_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.__component.addDocumentEventListener(self.events_listener_document_event)
        event.remove_callback = True

    def _on_print_job_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.__component.addPrintJobListener(self.events_listener_print_job)
        event.remove_callback = True

    def _on_close_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.__component.addCloseListener(self.events_listener_close)  # type: ignore
        event.remove_callback = True

    # endregion Lazy Listeners

    def get_selected_shapes(self) -> List[GenericShapeComp]:
        """Get the names of the selected shapes."""
        from ooodev.draw.shapes.draw_shape import DrawShape

        selection = self.get_selection()
        if selection is None:
            return []
        result = []
        if mInfo.Info.support_service(
            selection, "com.sun.star.drawing.Shapes", "com.sun.star.drawing.ShapeCollection"
        ):
            shapes = IndexAccessComp(selection)
            draw_shapes = [DrawShape(self, shape) for shape in shapes]  # type: ignore
            for draw_shape in draw_shapes:
                result.append(draw_shape.get_known_shape())
        return result
