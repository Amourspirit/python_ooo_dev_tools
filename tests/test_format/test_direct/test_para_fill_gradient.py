from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.para.area import Gradient, PresetGradientKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service


def test_write(loader, para_text) -> None:
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
        # p_len = len(para_text)

        # with Lo.ControllerLock():
        # using ControllerLock() is about 25% faster in this case
        # add a paragraph for each enum and test values
        presets = (
            PresetGradientKind.PASTEL_BOUQUET,
            PresetGradientKind.PASTEL_DREAM,
            PresetGradientKind.BLUE_TOUCH,
            PresetGradientKind.BLANK_GRAY,
            PresetGradientKind.SPOTTED_GRAY,
            PresetGradientKind.LONDON_MIST,
            PresetGradientKind.TEAL_BLUE,
            PresetGradientKind.MIDNIGHT,
            PresetGradientKind.DEEP_OCEAN,
            PresetGradientKind.SUBMARINE,
            PresetGradientKind.GREEN_GRASS,
            PresetGradientKind.NEON_LIGHT,
            PresetGradientKind.SUNSHINE,
            PresetGradientKind.PRESENT,
            PresetGradientKind.MAHOGANY,
        )
        for preset in presets:
            pg = Gradient.from_preset(preset)
            Write.append_para(cursor=cursor, text=para_text, styles=(pg,))
            cursor_p.gotoEnd(False)
            cursor_p.gotoPreviousParagraph(True)
            fp = cast("FillProperties", cursor_p.TextParagraph)
            pg_obj = Gradient.from_obj(fp)
            assert pg_obj == pg
            cursor_p.gotoEnd(False)

        # test applying directly to cursor.
        cursor_p.gotoEnd(False)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        pg = Gradient.from_preset(PresetGradientKind.DEEP_OCEAN)
        pg.apply(fp)
        for _ in range(3):
            Write.append_para(cursor=cursor, text=para_text)
        cursor_p.gotoEnd(False)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        pg_obj = Gradient.from_obj(fp)
        assert pg_obj == pg

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
