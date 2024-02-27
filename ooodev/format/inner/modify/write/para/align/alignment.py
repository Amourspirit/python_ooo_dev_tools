# region Import
from __future__ import annotations
from typing import Any, cast
import uno
from ooo.dyn.style.paragraph_adjust import ParagraphAdjust
from ooo.dyn.text.paragraph_vert_align import ParagraphVertAlignEnum

from ooodev.format.inner.direct.write.para.align.alignment import Alignment as InnerAlignment
from ooodev.format.inner.direct.write.para.align.alignment import LastLineKind
from ooodev.format.inner.direct.write.para.align.writing_mode import WritingMode
from ooodev.format.inner.modify.write.para.para_style_base_multi import ParaStyleBaseMulti
from ooodev.format.writer.style.para.kind.style_para_kind import StyleParaKind

# endregion Import


class Alignment(ParaStyleBaseMulti):
    """
    Paragraph Alignment

    Any properties starting with ``prop_`` set or get current instance values.

    .. seealso::

        - :ref:`help_writer_format_modify_para_alignment`

    .. versionadded:: 0.9.0
    """

    # region init

    def __init__(
        self,
        *,
        align: ParagraphAdjust | None = None,
        align_vert: ParagraphVertAlignEnum | None = None,
        txt_direction: WritingMode | None = None,
        align_last: LastLineKind | None = None,
        expand_single_word: bool | None = None,
        snap_to_grid: bool | None = None,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> None:
        """
        Constructor

        Args:
            align (ParagraphAdjust, optional): Determines horizontal alignment of a paragraph.
            align_vert (ParagraphVertAlignEnum, optional): Determines vertical alignment of a paragraph.
            align_last (LastLineKind, optional): Determines the adjustment of the last line.
            expand_single_word (bool, optional): Determines if single words are stretched.
                It is only valid if ``align`` and ``align_last`` are also valid.
            snap_to_grid (bool, optional): Determines snap to text grid (if active).
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Family Style. Default ``ParagraphStyles``.

        Returns:
            None:

        See Also:
            - :ref:`help_writer_format_modify_para_alignment`
        """

        direct = InnerAlignment(
            align=align,
            align_vert=align_vert,
            txt_direction=txt_direction,
            align_last=align_last,
            expand_single_word=expand_single_word,
            snap_to_grid=snap_to_grid,
        )
        super().__init__()
        self._style_name = str(style_name)
        self._style_family_name = style_family

        self._set_style("direct", direct, *direct.get_attrs())

    # endregion init

    @classmethod
    def from_style(
        cls,
        doc: Any,
        style_name: StyleParaKind | str = StyleParaKind.STANDARD,
        style_family: str = "ParagraphStyles",
    ) -> Alignment:
        """
        Gets instance from Document.

        Args:
            doc (Any): UNO Document Object.
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to.
                Default is Default Paragraph Style.
            style_family (str, optional): Style family. Default ``ParagraphStyles``.

        Returns:
            Alignment: ``Alignment`` instance from document properties.
        """
        inst = cls(style_name=style_name, style_family=style_family)
        direct = InnerAlignment.from_obj(inst.get_style_props(doc))
        inst._set_style("direct", direct, *direct.get_attrs())
        return inst

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleParaKind):
        self._style_name = str(value)

    @property
    def prop_inner(self) -> InnerAlignment:
        """Gets/Sets Inner Alignment instance"""
        try:
            return self._direct_inner
        except AttributeError:
            self._direct_inner = cast(InnerAlignment, self._get_style_inst("direct"))
        return self._direct_inner

    @prop_inner.setter
    def prop_inner(self, value: InnerAlignment) -> None:
        if not isinstance(value, InnerAlignment):
            raise TypeError(f'Expected type of InnerAlignment, got "{type(value).__name__}"')
        self._del_attribs("_direct_inner")
        self._set_style("direct", value, *value.get_attrs())
