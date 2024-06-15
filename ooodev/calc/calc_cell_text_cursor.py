from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno


from ooodev.adapter.text.text_cursor_comp import TextCursorComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.write.partial.text_cursor_partial import TextCursorPartial

# from ooodev.adapter.style.character_properties_partial import CharacterPropertiesPartial
# from ooodev.adapter.style.paragraph_properties_partial import ParagraphPropertiesPartial
from ooodev.calc.partial.calc_cell_prop_partial import CalcCellPropPartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial

if TYPE_CHECKING:
    from com.sun.star.text import XTextCursor
    from ooodev.calc.calc_cell import CalcCell
else:
    XSheetCellCursor = Any


class CalcCellTextCursor(
    LoInstPropsPartial,
    TextCursorComp,
    QiPartial,
    PropPartial,
    TextCursorPartial["CalcCellTextCursor"],
    ServicePartial,
    CalcCellPropPartial,
    CalcDocPropPartial,
    CalcSheetPropPartial,
    StylePartial,
    TheDictionaryPartial,
):
    def __init__(self, owner: CalcCell, cursor: XTextCursor, lo_inst: LoInst | None = None) -> None:
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        TextCursorComp.__init__(self, cursor)  # type: ignore
        QiPartial.__init__(self, component=cursor, lo_inst=self.lo_inst)  # type: ignore
        PropPartial.__init__(self, component=cursor, lo_inst=self.lo_inst)  # type: ignore
        TextCursorPartial.__init__(self, owner=self, component=self.component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=cursor, lo_inst=self.lo_inst)
        CalcCellPropPartial.__init__(self, obj=owner)
        CalcSheetPropPartial.__init__(self, obj=owner.calc_sheet)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)
        StylePartial.__init__(self, component=cursor)
        TheDictionaryPartial.__init__(self)
