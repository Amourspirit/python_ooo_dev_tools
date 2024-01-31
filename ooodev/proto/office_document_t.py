from __future__ import annotations
from typing import TYPE_CHECKING


from ooodev.utils.partial.service_partial_t import ServicePartialT
from ooodev.format.inner.style_partial_t import StylePartialT
from ooodev.utils.partial.lo_inst_props_partial_t import LoInstPropsPartialT
from ooodev.utils.partial.qi_partial_t import QiPartialT
from ooodev.utils.partial.gui_partial_t import GuiPartialT
from ooodev.dialog.partial.create_dialog_partial_t import CreateDialogPartialT
from ooodev.utils.inst.lo.doc_type import DocType

if TYPE_CHECKING:
    try:
        from typing import Protocol
    except ImportError:
        from typing_extensions import Protocol
else:
    Protocol = object


class OfficeDocumentT(
    LoInstPropsPartialT, QiPartialT, ServicePartialT, GuiPartialT, StylePartialT, CreateDialogPartialT, Protocol
):
    """Represents the common interface for all office documents."""

    DOC_TYPE: DocType
