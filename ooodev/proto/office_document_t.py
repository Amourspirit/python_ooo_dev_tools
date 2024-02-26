from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.dialog.partial.create_dialog_partial_t import CreateDialogPartialT
from ooodev.events.events_t import EventsT
from ooodev.format.inner.style_partial_t import StylePartialT
from ooodev.loader.inst.doc_type import DocType
from ooodev.utils.partial.dispatch_partial_t import DispatchPartialT
from ooodev.utils.partial.gui_partial_t import GuiPartialT
from ooodev.utils.partial.lo_inst_props_partial_t import LoInstPropsPartialT
from ooodev.utils.partial.qi_partial_t import QiPartialT
from ooodev.utils.partial.service_partial_t import ServicePartialT

if TYPE_CHECKING:
    from typing_extensions import Protocol
else:
    Protocol = object


class OfficeDocumentT(
    LoInstPropsPartialT,
    QiPartialT,
    ServicePartialT,
    GuiPartialT,
    StylePartialT,
    CreateDialogPartialT,
    EventsT,
    DispatchPartialT,
    Protocol,
):
    """Represents the common interface for all office documents."""

    DOC_TYPE: DocType
