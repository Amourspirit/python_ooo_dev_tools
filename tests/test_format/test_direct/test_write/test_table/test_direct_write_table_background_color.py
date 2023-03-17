from __future__ import annotations
import pytest
from typing import cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.table.background import Color
from ooodev.format import StandardColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.table_helper import TableHelper


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

        style = Color(StandardColor.LIME_LIGHT3)

        table = Write.add_table(cursor=cursor, table_data=tbl_data, first_row_header=False, styles=(style,))

        clr = Color.from_obj(table)
        assert clr.prop_color == StandardColor.LIME_LIGHT3

        style.empty.apply(table)
        clr = Color.from_obj(table)
        assert clr.prop_color == -1

        style = Color(color=StandardColor.INDIGO_LIGHT2)
        style.apply(table)
        clr = Color.from_obj(table)
        assert clr.prop_color == StandardColor.INDIGO_LIGHT2

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
