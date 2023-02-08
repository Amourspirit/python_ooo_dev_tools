from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno

from ooodev.format.direct.para.transparent.gradient import Gradient, GradientStyle

# from ooodev.format.writer.direct.para.transparency import Gradient, GradientStyle
# from ooodev.format.direct.para.area.color import Color
from ooodev.format.writer.direct.para.area import Color
from ooodev.format.direct.fill.transparent.gradient import FillTransparentGrad
from ooodev.format import Styler
from ooodev.format.writer.style.para import Para
from ooodev.format import StandardColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.draw import Draw
from ooodev.office.write import Write

if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def _test_write(loader, para_text) -> None:
    # Not currently working. Pararaph background is not transparent for gradients using this method.

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
        cursor_p = Write.get_paragraph_cursor(cursor)

        Write.append_para(cursor=cursor, text=para_text)
        cursor_p.gotoEnd(False)

        dc = Color(StandardColor.LIME)
        # ParaBackColor 2088883226  Color   Lime  Linear  Start 0%    End value 100%
        tp = Gradient(style=GradientStyle.LINEAR, angle=0, start_value=0, end_value=100)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc, tp))
        Para.default.apply(cursor)
        cursor_p.gotoEnd(False)
        cursor_p.gotoPreviousParagraph(False)
        cursor_p.gotoStartOfParagraph(False)
        cursor_p.gotoEndOfParagraph(True)
        pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        grad = cast(FillTransparentGrad, tp._get_style_inst("fill_grad")).get_gradient()

        cursor_p.gotoEnd(False)

        page = Write.get_draw_page(doc)
        rs = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
        dc = Color(StandardColor.DEFAULT_BLUE)
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
