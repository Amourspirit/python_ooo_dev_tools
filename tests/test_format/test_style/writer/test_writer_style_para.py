from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.inner.direct.write.para.align import Alignment, LastLineKind
from ooodev.format.writer.style import Para, StyleParaKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.color import CommonColor


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_write(loader, para_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        p_len = len(para_text)
        al = Alignment().align_right
        al.apply(cursor)
        Write.append_para(cursor=cursor, text=para_text)

        cursor.goLeft(1, False)
        cursor.gotoStart(True)
        pp = cast("ParagraphProperties", cursor)
        assert pp.ParaAdjust == 1  # ParagraphAdjust.RIGHT
        cursor.gotoEnd(False)
        Para.default.apply(cursor)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)
        assert pp.ParaAdjust == 0

        al = Alignment(snap_to_grid=False, align_last=LastLineKind.CENTER).justified
        al.apply(cursor)
        Write.append_para(cursor=cursor, text=para_text)
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaAdjust == 2  # ParagraphAdjust.BLOCK
        assert pp.ParaLastLineAdjust == LastLineKind.CENTER.value
        assert cursor.SnapToGrid == False
        cursor.gotoEnd(False)
        Para.default.apply(cursor)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)
        assert pp.ParaAdjust == 0
        assert pp.ParaLastLineAdjust == 0

        al = Alignment(align_last=LastLineKind.JUSTIFY, expand_single_word=True).justified
        al.apply(cursor)
        Write.append_para(cursor=cursor, text=para_text)
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaAdjust == 2  # ParagraphAdjust.BLOCK
        assert pp.ParaLastLineAdjust == LastLineKind.JUSTIFY.value
        assert pp.ParaExpandSingleWord == True
        cursor.gotoEnd(False)
        Para.default.apply(cursor)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)
        assert pp.ParaAdjust == 0
        assert pp.ParaLastLineAdjust == 0
        assert pp.ParaExpandSingleWord == False

        st = Para(name=StyleParaKind.TITLE)
        txt = "Title Text"
        t_len = len(txt)
        Write.append_para(cursor=cursor, text=txt, styles=(st,))
        cursor.goLeft(t_len + 1, False)
        cursor.goRight(t_len, True)
        assert pp.ParaStyleName == str(StyleParaKind.TITLE)
        cursor.gotoEnd(False)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)

        st = Para().h1
        txt = "H1 Text"
        t_len = len(txt)
        Write.append_para(cursor=cursor, text=txt, styles=(st,))
        cursor.goLeft(t_len + 1, False)
        cursor.goRight(t_len, True)
        assert pp.ParaStyleName == str(StyleParaKind.HEADING_1)
        cursor.gotoEnd(False)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)

        st = Para().h2
        txt = "H2 Text"
        t_len = len(txt)
        Write.append_para(cursor=cursor, text=txt, styles=(st,))
        cursor.goLeft(t_len + 1, False)
        cursor.goRight(t_len, True)
        assert pp.ParaStyleName == str(StyleParaKind.HEADING_2)
        cursor.gotoEnd(False)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)

        st = Para().h3
        txt = "H3 Text"
        t_len = len(txt)
        Write.append_para(cursor=cursor, text=txt, styles=(st,))
        cursor.goLeft(t_len + 1, False)
        cursor.goRight(t_len, True)
        assert pp.ParaStyleName == str(StyleParaKind.HEADING_3)
        cursor.gotoEnd(False)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)

        st = Para().h4
        txt = "H4 Text"
        t_len = len(txt)
        Write.append_para(cursor=cursor, text=txt, styles=(st,))
        cursor.goLeft(t_len + 1, False)
        cursor.goRight(t_len, True)
        assert pp.ParaStyleName == str(StyleParaKind.HEADING_4)
        cursor.gotoEnd(False)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)

        st = Para().h5
        txt = "H5 Text"
        t_len = len(txt)
        Write.append_para(cursor=cursor, text=txt, styles=(st,))
        cursor.goLeft(t_len + 1, False)
        cursor.goRight(t_len, True)
        assert pp.ParaStyleName == str(StyleParaKind.HEADING_5)
        cursor.gotoEnd(False)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)

        st = Para().h6
        txt = "H6 Text"
        t_len = len(txt)
        Write.append_para(cursor=cursor, text=txt, styles=(st,))
        cursor.goLeft(t_len + 1, False)
        cursor.goRight(t_len, True)
        assert pp.ParaStyleName == str(StyleParaKind.HEADING_6)
        cursor.gotoEnd(False)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)

        st = Para().h7
        txt = "H7 Text"
        t_len = len(txt)
        Write.append_para(cursor=cursor, text=txt, styles=(st,))
        cursor.goLeft(t_len + 1, False)
        cursor.goRight(t_len, True)
        assert pp.ParaStyleName == str(StyleParaKind.HEADING_7)
        cursor.gotoEnd(False)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)

        st = Para().h8
        txt = "H8 Text"
        t_len = len(txt)
        Write.append_para(cursor=cursor, text=txt, styles=(st,))
        cursor.goLeft(t_len + 1, False)
        cursor.goRight(t_len, True)
        assert pp.ParaStyleName == str(StyleParaKind.HEADING_8)
        cursor.gotoEnd(False)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)

        st = Para().h9
        txt = "H9 Text"
        t_len = len(txt)
        Write.append_para(cursor=cursor, text=txt, styles=(st,))
        cursor.goLeft(t_len + 1, False)
        cursor.goRight(t_len, True)
        assert pp.ParaStyleName == str(StyleParaKind.HEADING_9)
        cursor.gotoEnd(False)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)

        st = Para().h10
        txt = "H10 Text"
        t_len = len(txt)
        Write.append_para(cursor=cursor, text=txt, styles=(st,))
        cursor.goLeft(t_len + 1, False)
        cursor.goRight(t_len, True)
        assert pp.ParaStyleName == str(StyleParaKind.HEADING_10)
        cursor.gotoEnd(False)
        assert pp.ParaStyleName == str(StyleParaKind.STANDARD)

        style = Para(name=StyleParaKind.STANDARD)
        xprops = style.get_style_props()
        assert xprops is not None
        xprops.setPropertyValue("FillColor", CommonColor.CORAL)
        val = cast(int, xprops.getPropertyValue("FillColor"))
        assert val == CommonColor.CORAL

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
