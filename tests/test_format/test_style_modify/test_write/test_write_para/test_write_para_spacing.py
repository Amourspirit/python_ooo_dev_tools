from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.para.indent_space import Spacing
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


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
        Write.append_para(cursor=cursor, text=para_text)

        style = Spacing(above=4.5, below=3.5)
        style.apply(doc)
        props = style.get_style_props(doc)
        assert props.getPropertyValue("ParaTopMargin") in (448, 449, 450, 451, 452)
        assert props.getPropertyValue("ParaBottomMargin") in (348, 349, 350, 351, 352)

        f_style = Spacing.from_style(doc)
        assert f_style.prop_inner.prop_above.value == pytest.approx(4.5, rel=1e-2)
        assert f_style.prop_inner.prop_below.value == pytest.approx(3.5, rel=1e-2)
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
