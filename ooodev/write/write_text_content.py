from __future__ import annotations
from typing import TYPE_CHECKING, TypeVar, Generic
import uno

from ooodev.mock import mock_g
from ooodev.adapter.text.text_content_comp import TextContentComp
from ooodev.format.inner.style_partial import StylePartial
from ooodev.loader import lo as mLo
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.loader.inst.lo_inst import LoInst
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.write.partial.write_doc_prop_partial import WriteDocPropPartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial

if TYPE_CHECKING:
    from ooodev.proto.component_proto import ComponentT
    from com.sun.star.text import XTextContent
    from ooodev.adapter.style.character_properties_comp import CharacterPropertiesComp
    from ooodev.write.write_paragraph import WriteParagraph

T = TypeVar("T", bound="ComponentT")


class WriteTextContent(
    Generic[T],
    LoInstPropsPartial,
    TextContentComp,
    WriteDocPropPartial,
    QiPartial,
    ServicePartial,
    TheDictionaryPartial,
    StylePartial,
):
    """Represents writer text content."""

    def __init__(self, owner: T, component: XTextContent, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (T): Owner of this component.
            component (XTextContent): UNO object that supports ``com.sun.star.text.TextContent`` service.
            lo_inst (LoInst, optional): Lo instance. Defaults to ``None``.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        self._owner = owner
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        if not isinstance(owner, WriteDocPropPartial):
            raise TypeError("WriteDocPropPartial is not inherited by owner.")
        WriteDocPropPartial.__init__(self, obj=owner.write_doc)  # type: ignore
        TextContentComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)  # type: ignore
        StylePartial.__init__(self, component=component)
        TheDictionaryPartial.__init__(self)

    def get_character_properties(self) -> CharacterPropertiesComp | None:
        """
        Get character properties.

        Returns:
            CharacterPropertiesComp | None: Character properties component or ``None`` if not supported.
        """
        # pylint: disable=import-outside-toplevel
        if not self.support_service("com.sun.star.style.CharacterProperties"):
            return None
        from ooodev.adapter.style.character_properties_comp import CharacterPropertiesComp

        return CharacterPropertiesComp(component=self.component)  # type: ignore

    def get_paragraph_properties(self) -> WriteParagraph[WriteTextContent[T]] | None:
        """
        Get paragraph properties.

        Returns:
            WriteParagraph | None: Paragraph properties or ``None`` if not supported.
        """
        # pylint: disable=import-outside-toplevel
        if not self.support_service("com.sun.star.style.ParagraphProperties"):
            return None
        from ooodev.write.write_paragraph import WriteParagraph

        return WriteParagraph(owner=self, component=self.component, lo_inst=self.lo_inst)  # type: ignore

    def is_text_table(self) -> bool:
        """
        Check if this text content is a table.

        Returns:
            bool: ``True`` if component supports ``com.sun.star.text.TextTable`` service; Otherwise, ``False``.
        """
        return self.support_service("com.sun.star.text.TextTable")

    def is_text_frame(self) -> bool:
        """
        Check if this text content is a frame.

        Returns:
            bool: ``True`` if component supports ``com.sun.star.text.TextFrame`` service; Otherwise, ``False``.
        """
        return self.support_service("com.sun.star.text.TextTable")

    # region Properties
    @property
    def owner(self) -> T:
        """Component Owner"""
        return self._owner

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.adapter.style.character_properties_comp import CharacterPropertiesComp
    from ooodev.write.write_paragraph import WriteParagraph
