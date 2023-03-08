from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.table.properties import (
    TableProperties,
    TblAuto,
    TblWidth,
    TblLeft,
    TblRight,
    TblCenter,
    TblFromLeft,
    TblFromLeftWidth,
    TblManual,
    TblRelLeftByWidth,
    TblRelFromLeft,
    TblRelRightByWidth,
    TblRelCenter,
    TableAlignKind,
)
from ooodev.utils.color import StandardColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.table_helper import TableHelper
from ooodev.utils.data_type.unit_mm100 import UnitMM100


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

        above = UnitMM100.from_mm(2.0)
        below = UnitMM100.from_mm(1.8)
        style = TableProperties(name="My_Table", size=TblAuto(above=above, below=below))

        table = Write.add_table(cursor=cursor, table_data=tbl_data)
        style.apply(table)

        tp = TableProperties.from_obj(table)
        pobj = tp.prop_obj
        assert isinstance(pobj, TblAuto)
        assert tp.prop_name == "My_Table"
        assert pobj.prop_above.get_value_mm100() in range(above.value - 2, above.value + 3)  # +- 2
        assert pobj.prop_below.get_value_mm100() in range(below.value - 2, below.value + 3)  # +- 2

        margin = UnitMM100.from_mm(3.0)
        style = TableProperties(size=TblLeft(margin=margin, above=above, below=below))
        style.apply(table)

        tp = TableProperties.from_obj(table)
        pobj = tp.prop_obj
        assert isinstance(pobj, TblLeft)
        assert tp.prop_name == "My_Table"
        assert pobj.prop_above.get_value_mm100() in range(above.value - 2, above.value + 3)  # +- 2
        assert pobj.prop_below.get_value_mm100() in range(below.value - 2, below.value + 3)  # +- 2
        assert pobj.prop_margin.get_value_mm100() in range(margin.value - 2, margin.value + 3)  # +- 2

        width = UnitMM100.from_mm(150.0)
        style = TableProperties(size=TblWidth(width=width, above=above, below=below))
        style.apply(table)

        tp = TableProperties.from_obj(table)
        pobj = tp.prop_obj
        assert isinstance(pobj, TblWidth)
        assert tp.prop_name == "My_Table"
        assert pobj.prop_above.get_value_mm100() in range(above.value - 2, above.value + 3)  # +- 2
        assert pobj.prop_below.get_value_mm100() in range(below.value - 2, below.value + 3)  # +- 2
        assert pobj.prop_width.get_value_mm100() in range(width.value - 2, width.value + 3)  # +- 2

        margin = UnitMM100.from_mm(100.0)
        style = TableProperties(size=TblRight(margin=margin, above=above, below=below))
        style.apply(table)

        tp = TableProperties.from_obj(table)
        pobj = tp.prop_obj
        assert isinstance(pobj, TblRight)
        assert tp.prop_name == "My_Table"
        assert pobj.prop_above.get_value_mm100() in range(above.value - 2, above.value + 3)  # +- 2
        assert pobj.prop_below.get_value_mm100() in range(below.value - 2, below.value + 3)  # +- 2
        assert pobj.prop_margin.get_value_mm100() in range(margin.value - 2, margin.value + 3)  # +- 2

        margin = UnitMM100.from_mm(50)
        style = TableProperties(size=TblCenter(margin=margin, above=above, below=below))
        style.apply(table)

        tp = TableProperties.from_obj(table)
        pobj = tp.prop_obj
        # from_obj for TblCenter will return TblWidth
        assert isinstance(pobj, TblWidth)
        assert tp.prop_name == "My_Table"
        assert pobj.prop_above.get_value_mm100() in range(above.value - 2, above.value + 3)  # +- 2
        assert pobj.prop_below.get_value_mm100() in range(below.value - 2, below.value + 3)  # +- 2

        margin = UnitMM100.from_mm(3.5)
        style = TableProperties(size=TblFromLeft(margin=margin, above=above, below=below))
        style.apply(table)

        tp = TableProperties.from_obj(table)
        pobj = tp.prop_obj
        assert isinstance(pobj, TblFromLeft)
        assert tp.prop_name == "My_Table"
        assert pobj.prop_above.get_value_mm100() in range(above.value - 2, above.value + 3)  # +- 2
        assert pobj.prop_below.get_value_mm100() in range(below.value - 2, below.value + 3)  # +- 2
        assert pobj.prop_margin.get_value_mm100() in range(margin.value - 2, margin.value + 3)  # +- 2

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_write_abs(loader) -> None:
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

        above = UnitMM100.from_mm(2.0)
        below = UnitMM100.from_mm(1.8)
        style = TableProperties(name="My_Table", relative=False, align=TableAlignKind.AUTO)

        table = Write.add_table(cursor=cursor, table_data=tbl_data)
        style.apply(table)

        tp = TableProperties.from_obj(table)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
