from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno

from ooodev.adapter.drawing.text_comp import TextComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.office import draw as mDraw
from ooodev.office.partial.office_document_prop_partial import OfficeDocumentPropPartial
from ooodev.loader import lo as mLo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.draw import draw_text_cursor

if TYPE_CHECKING:
    from com.sun.star.text import XText
    from ooodev.proto.component_proto import ComponentT

_T = TypeVar("_T", bound="ComponentT")


class DrawText(
    Generic[_T],
    LoInstPropsPartial,
    OfficeDocumentPropPartial,
    TextComp,
    QiPartial,
    StylePartial,
    ServicePartial,
    TheDictionaryPartial,
):
    """
    Represents text content.

    Contains Enumeration Access.
    """

    def __init__(self, owner: _T, component: XText, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (_T): Owner of this component.
            component (XText): UNO object that supports ``com.sun.star.text.Text`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, OfficeDocumentPropPartial):
            raise ValueError("owner must be an instance of OfficeDocumentPropPartial")
        OfficeDocumentPropPartial.__init__(self, owner.office_doc)
        TextComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        TheDictionaryPartial.__init__(self)

    def add_bullet(self, level: int, text: str) -> None:
        """
        Add bullet text to the end of the bullets text area, specifying
        the nesting of the bullet using a numbering level value
        (numbering starts at 0).

        Args:
            level (int): Bullet Level
            text (str): Bullet Text

        Raises:
            DrawError: If error adding bullet.

        Returns:
            None:
        """
        mDraw.Draw.add_bullet(self.component, level, text)

    def get_cursor(self) -> draw_text_cursor.DrawTextCursor[_T]:
        """
        Get the cursor for this text.

        Returns:
            DrawTextCursor[_T]: Cursor for this text.
        """
        return draw_text_cursor.DrawTextCursor(
            owner=self.owner, component=self.component.createTextCursor(), lo_inst=self.lo_inst
        )

    # region Properties
    @property
    def owner(self) -> _T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
