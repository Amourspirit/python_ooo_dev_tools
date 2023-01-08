from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.styles.para.writing_mode import WritingMode, WritingMode2Enum
from ooodev.styles.para.alignment import Alignment, ParagraphAdjust, ParagraphVertAlignEnum, LastLineKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service
    from com.sun.star.style import ParagraphPropertiesComplex  # service


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
    wm = cast(WritingMode, al._get_style("txt_direction")[0])
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
    wm = cast(WritingMode, al._get_style("txt_direction")[0])
    assert wm.prop_mode == WritingMode2Enum.PAGE


def test_alignment_justify() -> None:
    al = cast(Alignment, Alignment.default)
    j = al.justified
    assert j.prop_align == ParagraphAdjust.BLOCK

    j = al.align_center
    assert j.prop_align == ParagraphAdjust.CENTER

    j = al.align_left
    assert j.prop_align == ParagraphAdjust.LEFT

    j = al.align_right
    assert j.prop_align == ParagraphAdjust.RIGHT


def test_alignment_copy() -> None:
    al = cast(Alignment, Alignment.default.copy())
    assert al.prop_align == ParagraphAdjust.LEFT
    assert al.prop_align_vert == ParagraphVertAlignEnum.AUTOMATIC
    assert al.prop_align_last == LastLineKind.START
    assert al.prop_expand_single_word == False
    assert al.prop_snap_to_grid == True
    wm = cast(WritingMode, al._get_style("txt_direction")[0])
    assert wm.prop_mode == WritingMode2Enum.PAGE


def test_alignemnt_write(loader) -> None:
    delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        p_txt = (
            "To Sherlock Holmes she is always THE woman. I have seldom heard"
            " him mention her under any other name. In his eyes she eclipses"
            " and predominates the whole of her sex. It was not that he felt"
            " any emotion akin to love for Irene Adler. All emotions, and that"
            " one particularly, were abhorrent to his cold, precise but"
            " admirably balanced mind. He was, I take it, the most perfect"
            " reasoning and observing machine that the world has seen, but as a"
            " lover he would have placed himself in a false position. He never"
            " spoke of the softer passions, save with a gibe and a sneer. They"
            " were admirable things for the observer--excellent for drawing the"
            " veil from men's motives and actions. But for the trained reasoner"
            " to admit such intrusions into his own delicate and finely"
            " adjusted temperament was to introduce a distracting factor which"
            " might throw a doubt upon all his mental results. Grit in a"
            " sensitive instrument, or a crack in one of his own high-power"
            " lenses, would not be more disturbing than a strong emotion in a"
            " nature such as his. And yet there was but one woman to him, and"
            " that woman was the late Irene Adler, of dubious and questionable memory."
        )

        p_len = len(p_txt)

        Write.append_para(cursor=cursor, text=p_txt, styles=(Alignment().align_right,))

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        assert pp.ParaAdjust == 1  # ParagraphAdjust.RIGHT
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=p_txt, styles=(Alignment().align_left,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaAdjust == 0  # ParagraphAdjust.LEFT
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=p_txt, styles=(Alignment().align_center,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaAdjust == 3  # ParagraphAdjust.CENTER
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=p_txt, styles=(Alignment().justified,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaAdjust == 2  # ParagraphAdjust.BLOCK
        cursor.gotoEnd(False)

        Write.append_para(
            cursor=cursor,
            text=p_txt,
            styles=(Alignment(snap_to_grid=False, align_last=LastLineKind.CENTER).justified,),
        )
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaAdjust == 2  # ParagraphAdjust.BLOCK
        assert pp.ParaLastLineAdjust == LastLineKind.CENTER.value
        assert cursor.SnapToGrid == False
        cursor.gotoEnd(False)

        Write.append_para(
            cursor=cursor,
            text=p_txt,
            styles=(Alignment(align_last=LastLineKind.START).justified,),
        )
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaAdjust == 2  # ParagraphAdjust.BLOCK
        assert pp.ParaLastLineAdjust == LastLineKind.START.value
        cursor.gotoEnd(False)

        Write.append_para(
            cursor=cursor,
            text=p_txt,
            styles=(Alignment(align_last=LastLineKind.JUSTIFY, expand_single_word=True).justified,),
        )
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaAdjust == 2  # ParagraphAdjust.BLOCK
        assert pp.ParaLastLineAdjust == LastLineKind.JUSTIFY.value
        assert pp.ParaExpandSingleWord == True
        cursor.gotoEnd(False)

        Write.append_para(
            cursor=cursor,
            text=p_txt,
            styles=(Alignment(txt_direction=WritingMode().bt_lr),),
        )
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ppc = cast("ParagraphPropertiesComplex", cursor)
        assert ppc.WritingMode == WritingMode2Enum.BT_LR.value
        cursor.gotoEnd(False)

        Write.append_para(
            cursor=cursor,
            text=p_txt,
            styles=(Alignment(txt_direction=WritingMode().tb_lr),),
        )
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ppc = cast("ParagraphPropertiesComplex", cursor)
        assert ppc.WritingMode == WritingMode2Enum.TB_LR.value
        cursor.gotoEnd(False)

        # reset text direction
        al = Alignment(txt_direction=WritingMode().lr_tb)
        al.apply_style(cursor)

        Write.append_para(cursor=cursor, text=p_txt)
        Write.style_prev_paragraph(cursor=cursor, styles=(Alignment().justified,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaAdjust == 2  # ParagraphAdjust.BLOCK
        cursor.gotoEnd(False)

        assert pp.ParaAdjust == 0  # ParagraphAdjust.LEFT

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
