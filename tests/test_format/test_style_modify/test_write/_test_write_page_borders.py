from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.borders import Borders, Side, StylePageKind, BorderLineKind
from ooodev.format import CommonColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


def _test_write(loader, para_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        Write.append_para(cursor=cursor, text=para_text)

        border = Borders(border_side=Side(line=BorderLineKind.DOUBLE, color=CommonColor.DARK_RED))
        border.apply(doc)

        oth = Borders.from_obj(doc, border.prop_style_name)
        assert oth == border

        border = Borders(
            border_side=Side(line=BorderLineKind.DOUBLE, color=CommonColor.DARK_RED),
            style_name=StylePageKind.FIRST_PAGE,
        )
        border.apply(doc)

        oth = Borders.from_obj(doc, border.prop_style_name)
        assert oth == border

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
