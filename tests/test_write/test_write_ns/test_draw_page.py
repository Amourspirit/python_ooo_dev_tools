import os
import pytest
from pathlib import Path
from typing import cast

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.lo import Lo
from ooodev.write import Write
from ooodev.write import WriteDoc


def test_draw_page(loader):
    # see comments Write.add_image_shape(cursor=cursor, fnm=im_fnm) Line: 181

    visible = False if Lo.bridge_connector.headless else True
    delay = 0  # 500
    doc = WriteDoc(Write.create_doc(loader))
    try:
        if visible:
            doc.set_visible(visible=visible)

        pages = doc.get_draw_pages()
        page = pages[0]
        rect = page.draw_rectangle(10, 10, 100, 100)
        assert rect is not None
        # GUI.set_visible(True, doc)
        Lo.delay(delay)

    finally:
        doc.close_doc()
