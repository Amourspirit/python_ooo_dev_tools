import os
import pytest
from pathlib import Path
from typing import cast

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.lo import Lo
from ooodev.calc import Calc, CalcDoc


def test_draw_page(loader):
    visible = False if Lo.bridge_connector.headless else True
    delay = 0  # 500
    doc = CalcDoc(Calc.create_doc(loader))
    try:
        if visible:
            doc.set_visible(visible=visible)
        sheet = doc.sheets[0]

        rect = sheet.draw_page.draw_rectangle(10, 10, 100, 100)
        assert rect is not None
        assert len(sheet.draw_page) == 1

        pg = doc.draw_pages[0]
        assert len(pg) == 1
        assert pg is not None
        Lo.delay(delay)

    finally:
        doc.close_doc()
