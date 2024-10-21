from __future__ import annotations
import pytest
from typing import cast, TYPE_CHECKING

if __name__ == "__main__":
    pytest.main([__file__])

from ooo.dyn.style.graphic_location import GraphicLocation

from ooodev.format.writer.direct.table.background import (
    Img,
    PresetImageKind,
)
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.table_helper import TableHelper

if TYPE_CHECKING:
    from com.sun.star.text import TextTable


def test_write(loader) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        tbl_data = TableHelper.make_2d_array(num_rows=5, num_cols=5)

        style = Img.from_preset(PresetImageKind.PAPER_TEXTURE)

        table = Write.add_table(cursor=cursor, table_data=tbl_data, first_row_header=False)
        style.apply(table)
        tt = cast("TextTable", table)
        assert tt.BackColor == -1
        assert tt.BackGraphic is not None
        assert tt.BackTransparent
        assert tt.BackGraphicLocation == GraphicLocation.TILED

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
