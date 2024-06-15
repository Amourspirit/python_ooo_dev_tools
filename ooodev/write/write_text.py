from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING, TypeVar, Generic, overload
import contextlib
import uno
from com.sun.star.text import XTextRange

from ooodev.adapter.text.relative_text_content_insert_partial import RelativeTextContentInsertPartial
from ooodev.adapter.text.text_comp import TextComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.utils import info as mInfo
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.write import write_paragraphs as mWriteParagraphs
from ooodev.write.write_text_content import WriteTextContent

if TYPE_CHECKING:
    from com.sun.star.text import XText
    from com.sun.star.text import XTextRange
    from com.sun.star.text import XTextContent
    from ooodev.proto.component_proto import ComponentT
    from ooodev.write.write_text_cursor import WriteTextCursor
    from ooodev.write.write_text_range import WriteTextRange

T = TypeVar("T", bound="ComponentT")


class WriteText(
    LoInstPropsPartial,
    WriteDocPropPartial,
    TextComp,
    RelativeTextContentInsertPartial,
    QiPartial,
    TheDictionaryPartial,
    StylePartial,
    Generic[T],
):
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
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        TextComp.__init__(self, component)  # type: ignore
        RelativeTextContentInsertPartial.__init__(self, component=component, interface=None)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        TheDictionaryPartial.__init__(self)
        StylePartial.__init__(self, component=component)

    def get_paragraphs(self) -> mWriteParagraphs.WriteParagraphs[T]:
        """Returns the paragraphs of this text."""
        return mWriteParagraphs.WriteParagraphs(owner=self.owner, component=self.component, lo_inst=self.lo_inst)

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
            rng = self.lo_inst.qi(XTextRange, self.owner.component)
        if rng is None:
            raise TypeError("owner must be XTextRange when rng is None")
        TextComp.insert_text_content(self, rng, content, absorb)

    # region SimpleTextPartial overrides
    def create_text_cursor(self) -> WriteTextCursor[WriteText]:
        """
        Creates a new text cursor.

        Returns:
            WriteTextCursor[WriteText]: The new text cursor.
        """
        # pylint: disable=import-outside-toplevel
        from .write_text_cursor import WriteTextCursor

        cursor = self.component.createTextCursor()
        return WriteTextCursor(owner=self, component=cursor, lo_inst=self.lo_inst)

    def create_text_cursor_by_range(self, text_position: WriteTextRange | XTextRange) -> WriteTextCursor[WriteText]:
        """
        The initial position is set to ``text_position``.

        Args:
            text_position (WriteTextRange, XTextRange): The initial position of the new text cursor.

        Returns:
            WriteTextCursor[WriteText]: The new text cursor.
        """
        # pylint: disable=import-outside-toplevel
        from .write_text_cursor import WriteTextCursor
        from .write_text_range import WriteTextRange

        if mInfo.Info.is_instance(text_position, WriteTextRange):
            rng = cast("XTextRange", text_position.component)
        else:
            rng = cast("XTextRange", text_position)

        cursor = self.component.createTextCursorByRange(rng)
        return WriteTextCursor(owner=self, component=cursor, lo_inst=self.lo_inst)

    # endregion SimpleTextPartial overrides

    # region EnumerationAccessPartial overrides
    def _is_next_element_valid(self, element: Any) -> bool:
        """
        Gets if the next element is valid.
        This method is called when iterating over the elements of this class.

        Args:
            element (Any): Element

        Returns:
            bool: True in this class but can be overridden in child classes.
        """
        with contextlib.suppress(Exception):
            return element.supportsService("com.sun.star.text.TextContent")
        return False

    def __next__(self) -> WriteTextContent[WriteText[T]]:
        """
        Gets the next element.

        Returns:
            WriteTextContent[WriteText[T]]: Next element.
        """
        result = super().__next__()
        return WriteTextContent(owner=self, component=result, lo_inst=self.lo_inst)

    # endregion EnumerationAccessPartial overrides

    # region Properties
    @property
    def owner(self) -> T:
        """Owner of this component."""
        return self._owner

    # endregion Properties
