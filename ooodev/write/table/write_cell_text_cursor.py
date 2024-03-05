from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno


from ooodev.adapter.text.text_cursor_comp import TextCursorComp
from ooodev.adapter.style.character_properties_partial import CharacterPropertiesPartial
from ooodev.adapter.style.paragraph_properties_partial import ParagraphPropertiesPartial
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.write.partial.text_cursor_partial import TextCursorPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write.table.partial.write_table_prop_partial import WriteTablePropPartial

if TYPE_CHECKING:
    from com.sun.star.text import XTextCursor
    from ooodev.proto.component_proto import ComponentT


class WriteCellTextCursor(
    LoInstPropsPartial,
    WriteDocPropPartial,
    WriteTablePropPartial,
    TextCursorComp,
    QiPartial,
    PropPartial,
    TextCursorPartial["WriteCellTextCursor"],
    CharacterPropertiesPartial,
    ParagraphPropertiesPartial,
    ServicePartial,
    StylePartial,
):
    """Represents writer table text cursor."""

    def __init__(self, owner: ComponentT, cursor: XTextCursor, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        if not isinstance(owner, WriteTablePropPartial):
            raise TypeError("Owner must be an instance of WriteTablePropPartial.")

        WriteDocPropPartial.__init__(self, obj=owner.write_table.write_doc)
        WriteTablePropPartial.__init__(self, obj=owner.write_table)
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        TextCursorComp.__init__(self, cursor)  # type: ignore
        QiPartial.__init__(self, component=cursor, lo_inst=self.lo_inst)  # type: ignore
        PropPartial.__init__(self, component=cursor, lo_inst=self.lo_inst)  # type: ignore
        TextCursorPartial.__init__(self, owner=self, component=self.component, lo_inst=self.lo_inst)
        CharacterPropertiesPartial.__init__(self, component=cursor)  # type: ignore
        ParagraphPropertiesPartial.__init__(self, component=cursor)  # type: ignore
        ServicePartial.__init__(self, component=cursor, lo_inst=self.lo_inst)
        StylePartial.__init__(self, component=cursor)

    @property
    def owner(self) -> ComponentT:
        """Owner of this component."""
        return self._owner
