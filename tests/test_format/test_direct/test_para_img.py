from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.format.direct.para.area.img import (
    Img,
    ImageKind,
    ImgStyleKind,
    SizeMM,
    SizePercent,
    Offset,
    OffsetColumn,
    OffsetRow,
    RectanglePoint,
)

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service


def test_write(loader, para_text) -> None:
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        cursor = Write.get_cursor(doc)

        presets = (
            ImageKind.PAINTED_WHITE,
            ImageKind.PAPER_TEXTURE,
            ImageKind.PAPER_CRUMPLED,
        )
        cursor_p = Write.get_paragraph_cursor(cursor)
        for preset in presets:
            img = Img.from_preset(preset)

            Write.append_para(cursor=cursor, text=para_text, styles=(img,))

            cursor_p.gotoPreviousParagraph(True)
            fp = cast("FillProperties", cursor_p.TextParagraph)
            cursor_p.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
