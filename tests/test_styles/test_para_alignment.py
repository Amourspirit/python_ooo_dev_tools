from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.styles.para.writing_mode import WritingMode, WritingMode2Enum
from ooodev.styles.para.alignment import Alignment, ParagraphAdjust, ParagraphVertAlignEnum, LastLineKind
from ooodev.styles import CommonColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.styles.char.font import (
    Font,
    FontUnderlineEnum,
    FontFamilyEnum,
)

if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


def test_writing_mode_props() -> None:
    mode = WritingMode(mode=WritingMode2Enum.BT_LR)
    assert mode.prop_mode == WritingMode2Enum.BT_LR
    mode.prop_mode = WritingMode2Enum.CONTEXT
    assert mode.prop_mode == WritingMode2Enum.CONTEXT

    mode = WritingMode()
    assert mode.prop_mode is None

    assert WritingMode.default.prop_mode == WritingMode2Enum.PAGE


def test_alignment_props() -> None:
    al = Alignment()
    assert al.prop_align is None
    assert al.prop_align_last is None
    assert al.prop_align_vert is None
    assert al.prop_expand_single_word is None
    assert al.prop_snap_to_grid is None

    al = Alignment(align=ParagraphAdjust.BLOCK)
    assert al.prop_align == ParagraphAdjust.BLOCK
    al.prop_align = ParagraphAdjust.CENTER
    assert al.prop_align == ParagraphAdjust.CENTER
    al.prop_align = None
    assert al.prop_align is None

    al = Alignment(align_vert=ParagraphVertAlignEnum.BOTTOM)
    assert al.prop_align_vert == ParagraphVertAlignEnum.BOTTOM
    al.prop_align_vert = ParagraphVertAlignEnum.CENTER
    assert al.prop_align_vert == ParagraphVertAlignEnum.CENTER
    al.prop_align_vert = None
    assert al.prop_align_vert is None

    al = Alignment(txt_direction=WritingMode(WritingMode2Enum.PAGE))
    wm = cast(WritingMode, al._get_style("txt_direction"))
    assert wm.prop_mode == WritingMode2Enum.PAGE

    al = Alignment(align_last=LastLineKind.JUSTIFY)
    assert al.prop_align_last == LastLineKind.JUSTIFY
    al.prop_align_last = LastLineKind.CENTER
    assert al.prop_align_last == LastLineKind.CENTER
    al.prop_align_last = None
    assert al.prop_align_last is None

    al = Alignment(expand_single_word=True)
    assert al.prop_expand_single_word
    al.prop_expand_single_word = False
    assert al.prop_expand_single_word == False
    al.prop_expand_single_word = None
    assert al.prop_expand_single_word is None

    al = Alignment(snap_to_grid=True)
    assert al.prop_snap_to_grid
    al.prop_snap_to_grid = False
    assert al.prop_snap_to_grid == False
    al.prop_snap_to_grid = None
    assert al.prop_snap_to_grid is None


def test_alignment_default() -> None:
    al = cast(Alignment, Alignment.default)
    assert al.prop_align == ParagraphAdjust.LEFT
    assert al.prop_align_vert == ParagraphVertAlignEnum.AUTOMATIC
    assert al.prop_align_last == LastLineKind.START
    assert al.prop_expand_single_word == False
    assert al.prop_snap_to_grid == True
    wm = cast(WritingMode, al._get_style("txt_direction"))
    assert wm.prop_mode == WritingMode2Enum.PAGE


def test_alignment_copy() -> None:
    al = cast(Alignment, Alignment.default.copy())
    assert al.prop_align == ParagraphAdjust.LEFT
    assert al.prop_align_vert == ParagraphVertAlignEnum.AUTOMATIC
    assert al.prop_align_last == LastLineKind.START
    assert al.prop_expand_single_word == False
    assert al.prop_snap_to_grid == True
    wm = cast(WritingMode, al._get_style("txt_direction"))
    assert wm.prop_mode == WritingMode2Enum.PAGE
