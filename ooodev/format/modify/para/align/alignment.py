from __future__ import annotations
import uno
from ....direct.para.align.alignment import Alignment as DirectAlignment, LastLineKind as LastLineKind
from ....writer.style.para.kind import StyleParaKind as StyleParaKind
from ....direct.para.align.writing_mode import WritingMode as WritingMode
from ..para_style_base_multi import ParaStyleBaseMulti

from ooo.dyn.style.paragraph_adjust import ParagraphAdjust as ParagraphAdjust
from ooo.dyn.text.paragraph_vert_align import ParagraphVertAlignEnum as ParagraphVertAlignEnum


class Alignment(ParaStyleBaseMulti):
    """
    Paragraph Alignment

    Any properties starting with ``prop_`` set or get current instance values.

    .. versionadded:: 0.9.0
    """

    _DEFAULT = None

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
    ) -> None:
        """
        Constructor

        Args:
            align (ParagraphAdjust, optional): Determines horizontal alignment of a paragraph.
            align_vert (ParagraphVertAlignEnum, optional): Determines verticial alignment of a paragraph.
            align_last (LastLineKind, optional): Determines the adjustment of the last line.
            expand_single_word (bool, optional): Determines if single words are stretched.
                It is only valid if ``align`` and ``align_last`` are also valid.
            snap_to_grid (bool, optional): Determines snap to text grid (if active).
            style_name (StyleParaKind, str, optional): Specifies the Paragraph Style that instance applies to. Deftult is Default Paragraph Style.

        Returns:
            None:
        """

        dl = DirectAlignment(
            align=align,
            align_vert=align_vert,
            txt_direction=txt_direction,
            align_last=align_last,
            expand_single_word=expand_single_word,
            snap_to_grid=snap_to_grid,
        )
        self._style_name = str(style_name)
        super().__init__()
        self._set_style("fill_style", dl, *dl.get_attrs())

    @property
    def prop_style_name(self) -> str:
        """Gets/Sets property Style Name"""
        return self._style_name

    @prop_style_name.setter
    def prop_style_name(self, value: str | StyleParaKind):
        self._style_name = str(value)

    # endregion init
