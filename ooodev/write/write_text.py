from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic, overload
import uno
from com.sun.star.text import XTextRange

if TYPE_CHECKING:
    from com.sun.star.text import XText
    from com.sun.star.text import XTextContent

from ooodev.adapter.text.relative_text_content_insert_partial import RelativeTextContentInsertPartial
from ooodev.adapter.text.text_comp import TextComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.proto.component_proto import ComponentT
from ooodev.utils import lo as mLo
from ooodev.utils.inst.lo.lo_inst import LoInst
from ooodev.utils.partial.qi_partial import QiPartial
from . import write_paragraphs as mWriteParagraphs
from . import write_text_tables as mWriteTextTables

T = TypeVar("T", bound="ComponentT")


class WriteText(Generic[T], TextComp, RelativeTextContentInsertPartial, QiPartial, StylePartial):
    """
    Represents writer text content.

    Contains Enumeration Access.
    """

    def __init__(self, owner: T, component: XText, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XText): UNO object that supports ``com.sun.star.text.Text`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            self._lo_inst = mLo.Lo.current_lo
        else:
            self._lo_inst = lo_inst
        self._owner = owner
        TextComp.__init__(self, component)  # type: ignore
        RelativeTextContentInsertPartial.__init__(self, component=component, interface=None)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self._lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)

    def get_paragraphs(self) -> mWriteParagraphs.WriteParagraphs[T]:
        """Returns the paragraphs of this text."""
        return mWriteParagraphs.WriteParagraphs(owner=self.owner, component=self.component, lo_inst=self._lo_inst)

    def get_text_tables(self) -> mWriteTextTables.WriteTextTables[T]:
        """Returns the text tables of this text."""
        return mWriteTextTables.WriteTextTables(owner=self.owner, component=self.component, lo_inst=self._lo_inst)

    @overload
    def insert_text_content(self, content: XTextContent, absorb: bool) -> None:
        """
        Inserts a content, such as a text table, text frame or text field.

        Args:
            content (XTextContent): The content to be inserted.
            absorb (bool): Specifies whether the text spanned by xRange will be replaced.
                If ``True`` then the content of range will be replaced by content,
                otherwise content will be inserted at the end of xRange.

        Returns:
            None:
        """
        ...

    @overload
    def insert_text_content(
        self,
        content: XTextContent,
        absorb: bool,
        rng: XTextRange,
    ) -> None:
        """
        Inserts a content, such as a text table, text frame or text field.

        Args:
            content (XTextContent): The content to be inserted.
            absorb (bool): Specifies whether the text spanned by xRange will be replaced.
                If ``True`` then the content of range will be replaced by content,
                otherwise content will be inserted at the end of xRange.
            rng (XTextRange): The position at which the content is inserted.

        Returns:
            None:
        """
        ...

    def insert_text_content(self, content: XTextContent, absorb: bool, rng: XTextRange | None = None) -> None:
        """
        Inserts a content, such as a text table, text frame or text field.

        Args:
            rng (XTextRange): The position at which the content is inserted.
            content (XTextContent): The content to be inserted.
            absorb (bool): Specifies whether the text spanned by xRange will be replaced.
                If ``True`` then the content of range will be replaced by content,
                otherwise content will be inserted at the end of xRange.

        Returns:
            None:
        """
        if rng is None:
            rng = mLo.Lo.qi(XTextRange, self.owner.component)
            if rng is None:
                raise TypeError("owner must be XTextRange when rng is None")
        TextComp.insert_text_content(self, rng, content, absorb)

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
