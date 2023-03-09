from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.table.properties import (
    TableProperties,
    TblAuto,
    TblCenterWidth,
    TblLeft,
    TblRight,
    TblCenterLeft,
    TblFromLeft,
    TblFromLeftWidth,
    TblManualLeftRight,
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

        style = TableProperties(name="My_Table", relative=False, align=TableAlignKind.AUTO)

        table = Write.add_table(cursor=cursor, table_data=tbl_data)
        style.apply(table)

        tp = TableProperties.from_obj(table)
        assert tp.prop_left.value == 0
        assert tp.prop_right.value == 0
        assert tp.prop_width.value > 0

        # ==================== Center Width ===============================

        above100 = UnitMM100.from_mm(2.0)
        below100 = UnitMM100.from_mm(1.8)
        width100 = UnitMM100.from_mm(60)
        style = TableProperties(
            relative=False,
            align=TableAlignKind.CENTER,
            above=above100,
            below=below100,
            width=width100,
        )

        style.apply(table)

        tp = TableProperties.from_obj(table)
        assert tp.prop_left.value > 0
        assert tp.prop_right.value > 0
        assert tp.prop_right.get_value_mm100() in range(
            tp.prop_left.get_value_mm100() - 2, tp.prop_left.get_value_mm100() + 3
        )  # +- 2
        assert tp.prop_width.get_value_mm100() in range(width100.value - 2, width100.value + 3)
        assert tp.prop_above.get_value_mm100() in range(above100.value - 2, above100.value + 3)  # +- 2
        assert tp.prop_below.get_value_mm100() in range(below100.value - 2, below100.value + 3)  # +- 2

        # ==================== Center Left ===============================

        above100 = UnitMM100.from_mm(2.0)
        below100 = UnitMM100.from_mm(1.8)
        left100 = UnitMM100.from_mm(40)
        style = TableProperties(
            relative=False,
            align=TableAlignKind.CENTER,
            above=above100,
            below=below100,
            left=left100,
        )

        style.apply(table)

        tp = TableProperties.from_obj(table)
        assert tp.prop_width.value > 0
        assert tp.prop_left.get_value_mm100() in range(left100.value - 2, left100.value + 3)  # +- 2
        assert tp.prop_right.get_value_mm100() in range(left100.value - 2, left100.value + 3)  # +- 2
        assert tp.prop_above.get_value_mm100() in range(above100.value - 2, above100.value + 3)  # +- 2
        assert tp.prop_below.get_value_mm100() in range(below100.value - 2, below100.value + 3)  # +- 2

        # ===================== From Left Width ==============================

        above100 = UnitMM100.from_mm(2.0)
        below100 = UnitMM100.from_mm(1.8)
        width100 = UnitMM100.from_mm(60)
        style = TableProperties(
            relative=False,
            align=TableAlignKind.FROM_LEFT,
            above=above100,
            below=below100,
            width=width100,
        )

        style.apply(table)

        tp = TableProperties.from_obj(table)
        assert tp.prop_left.value > 0
        assert tp.prop_right.value == 0
        assert tp.prop_width.get_value_mm100() in range(width100.value - 2, width100.value + 3)

        # =====================From Left Left==============================

        left100 = UnitMM100.from_mm(55)
        style = TableProperties(
            relative=False,
            align=TableAlignKind.FROM_LEFT,
            above=above100,
            below=below100,
            left=left100,
        )

        style.apply(table)

        tp = TableProperties.from_obj(table)
        assert tp.prop_left.get_value_mm100() in range(left100.value - 2, left100.value + 3)
        assert tp.prop_width.value > 0

        # ===================== Left Width ==============================

        above100 = UnitMM100.from_mm(2.3)
        below100 = UnitMM100.from_mm(2.5)
        width100 = UnitMM100.from_mm(60)
        style = TableProperties(
            relative=False,
            align=TableAlignKind.LEFT,
            above=above100,
            below=below100,
            width=width100,
        )

        style.apply(table)

        tp = TableProperties.from_obj(table)
        assert tp.prop_left.value == 0
        assert tp.prop_right.value > 0
        assert tp.prop_width.get_value_mm100() in range(width100.value - 2, width100.value + 3)
        assert tp.prop_above.get_value_mm100() in range(above100.value - 2, above100.value + 3)  # +- 2
        assert tp.prop_below.get_value_mm100() in range(below100.value - 2, below100.value + 3)  # +- 2

        # =================== Left Right ================================

        above100 = UnitMM100.from_mm(2.3)
        below100 = UnitMM100.from_mm(2.5)
        right100 = UnitMM100.from_mm(60)
        style = TableProperties(
            relative=False,
            align=TableAlignKind.LEFT,
            above=above100,
            below=below100,
            right=right100,
        )

        style.apply(table)

        tp = TableProperties.from_obj(table)
        assert tp.prop_left.value == 0
        assert tp.prop_right.get_value_mm100() in range(right100.value - 2, right100.value + 3)
        assert tp.prop_above.get_value_mm100() in range(above100.value - 2, above100.value + 3)  # +- 2
        assert tp.prop_below.get_value_mm100() in range(below100.value - 2, below100.value + 3)  # +- 2

        # =================== Right Width ================================

        above100 = UnitMM100.from_mm(2.3)
        below100 = UnitMM100.from_mm(2.5)
        width100 = UnitMM100.from_mm(60)
        style = TableProperties(
            relative=False,
            align=TableAlignKind.RIGHT,
            above=above100,
            below=below100,
            width=width100,
        )

        style.apply(table)

        tp = TableProperties.from_obj(table)
        assert tp.prop_right.value == 0
        assert tp.prop_left.value > 0
        assert tp.prop_width.get_value_mm100() in range(width100.value - 2, width100.value + 3)
        assert tp.prop_above.get_value_mm100() in range(above100.value - 2, above100.value + 3)  # +- 2
        assert tp.prop_below.get_value_mm100() in range(below100.value - 2, below100.value + 3)  # +- 2

        # ==================== Right Left ===============================

        above100 = UnitMM100.from_mm(2.3)
        below100 = UnitMM100.from_mm(2.5)
        left100 = UnitMM100.from_mm(60)
        style = TableProperties(
            relative=False,
            align=TableAlignKind.RIGHT,
            above=above100,
            below=below100,
            left=left100,
        )

        style.apply(table)

        tp = TableProperties.from_obj(table)
        assert tp.prop_right.value == 0
        assert tp.prop_left.get_value_mm100() in range(left100.value - 2, left100.value + 3)
        assert tp.prop_above.get_value_mm100() in range(above100.value - 2, above100.value + 3)  # +- 2
        assert tp.prop_below.get_value_mm100() in range(below100.value - 2, below100.value + 3)  # +- 2

        # ==================== Manual Width ===============================

        above100 = UnitMM100.from_mm(2.3)
        below100 = UnitMM100.from_mm(2.5)
        width100 = UnitMM100.from_mm(60)
        style = TableProperties(
            relative=False,
            align=TableAlignKind.MANUAL,
            above=above100,
            below=below100,
            width=width100,
        )

        style.apply(table)

        tp = TableProperties.from_obj(table)
        assert tp.prop_width.get_value_mm100() in range(width100.value - 2, width100.value + 3)
        assert tp.prop_above.get_value_mm100() in range(above100.value - 2, above100.value + 3)  # +- 2
        assert tp.prop_below.get_value_mm100() in range(below100.value - 2, below100.value + 3)  # +- 2

        # ==================== Manual Left Right ===============================

        left100 = UnitMM100.from_mm(66)
        right100 = UnitMM100.from_mm(55)
        style = TableProperties(relative=False, align=TableAlignKind.MANUAL, left=left100, right=right100)

        style.apply(table)

        tp = TableProperties.from_obj(table)
        assert tp.prop_left.get_value_mm100() in range(left100.value - 2, left100.value + 3)  # +- 2
        assert tp.prop_right.get_value_mm100() in range(right100.value - 2, right100.value + 3)  # +- 2

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
