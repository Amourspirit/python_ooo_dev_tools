from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.para.transparency import Gradient, GradientStyle
from ooodev.format.direct.shared.transparent.gradient import FillTransparendGrad
from ooodev.format.direct.para.area.color import Color
from ooodev.format import Styler
from ooodev.format.writer.style.para import Para
from ooodev.format import CommonColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.draw import Draw
from ooodev.office.write import Write

if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_write(loader, para_text) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        Write.append_para(cursor=cursor, text=para_text)
        cursor_p = Write.get_paragraph_cursor(cursor)
        pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        dc = Color(CommonColor.LIME_GREEN)
        tp = Gradient(style=GradientStyle.LINEAR, angle=25, start_value=25, end_value=100)

        grad = cast(FillTransparendGrad, tp._get_style("fill_style")[0]).get_gradient()
        # add_gradient_to_table("Transparency 1", grad)

        # dc.apply(pp)
        # tp.apply(pp)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc, tp))
        # add_gradient_to_table("Transparency 1", grad)
        Para.default.apply(cursor)
        page = Write.get_draw_page(doc)
        rs = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
        dc = Color(CommonColor.LIGHT_BLUE)
        Styler.apply(rs, dc, tp)
        page.add(rs)
        # add_gradient_to_table("", None)
        # cursor_p = Write.get_paragraph_cursor(cursor)
        # cursor_p.gotoEnd(False)
        # cursor_p.gotoPreviousParagraph(True)
        # pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        # assert pp.DropCapCharStyleName == ""
        # assert pp.DropCapWholeWord == False
        # inner_dc = cast(DropCapStruct, dc._get_style("drop_cap")[0])
        # assert inner_dc == pp.DropCapFormat
        # cursor_p.gotoEnd(False)

        # dc = DropCaps(count=1, style=StyleCharKind.DROP_CAPS)
        # Write.append_para(cursor=cursor, text=para_text, styles=(dc,))
        # cursor_p.gotoEnd(False)
        # cursor_p.gotoPreviousParagraph(True)
        # pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        # assert pp.DropCapCharStyleName == StyleCharKind.DROP_CAPS.value
        # assert pp.DropCapWholeWord == False
        # inner_dc = cast(DropCapStruct, dc._get_style("drop_cap")[0])
        # assert inner_dc == pp.DropCapFormat
        # cursor_p.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
