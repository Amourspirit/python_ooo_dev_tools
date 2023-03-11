from __future__ import annotations
import pytest
from typing import cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.table.borders import (
    Borders,
    Shadow,
    Side,
    BorderLineKind,
    ShadowLocation,
    Padding,
    LineSize,
)
from ooodev.format import StandardColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.table_helper import TableHelper
from ooodev.utils.data_type.unit_mm100 import UnitMM100


def test_write_sides(loader) -> None:
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

        style = Borders(border_side=Side())

        table = Write.add_table(cursor=cursor, table_data=tbl_data, first_row_header=False, styles=(style,))
        # style.apply(table)

        tp = Borders.from_obj(table)
        assert tp.prop_inner_border_table is not None
        assert tp.prop_inner_shadow is not None
        assert tp.prop_inner_padding is not None

        side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.GREEN_DARK3, width=LineSize.THICK)
        style = style.fmt_left(side)
        style.apply(table)

        tp = Borders.from_obj(table)
        assert tp.prop_inner_border_table is not None
        assert tp.prop_inner_shadow is not None
        assert tp.prop_inner_padding is not None
        assert tp.prop_inner_border_table.prop_left.prop_color == StandardColor.GREEN_DARK3
        assert tp.prop_inner_border_table.prop_left.prop_line == BorderLineKind.DOUBLE

        side = Side(line=BorderLineKind.DOTTED, color=StandardColor.BRICK_LIGHT2, width=LineSize.EXTRA_THICK)
        style = style.fmt_right(side)
        style.apply(table)

        tp = Borders.from_obj(table)
        assert tp.prop_inner_border_table is not None
        assert tp.prop_inner_shadow is not None
        assert tp.prop_inner_padding is not None
        assert tp.prop_inner_border_table.prop_left.prop_color == StandardColor.GREEN_DARK3
        assert tp.prop_inner_border_table.prop_left.prop_line == BorderLineKind.DOUBLE
        assert tp.prop_inner_border_table.prop_right.prop_color == StandardColor.BRICK_LIGHT2
        assert tp.prop_inner_border_table.prop_right.prop_line == BorderLineKind.DOTTED

        side = Side(line=BorderLineKind.SOLID, color=StandardColor.BLUE_LIGHT2, width=LineSize.MEDIUM)
        style = style.fmt_bottom(side).fmt_top(side)
        style.apply(table)

        tp = Borders.from_obj(table)
        assert tp.prop_inner_border_table is not None
        assert tp.prop_inner_shadow is not None
        assert tp.prop_inner_padding is not None
        assert tp.prop_inner_border_table.prop_left.prop_color == StandardColor.GREEN_DARK3
        assert tp.prop_inner_border_table.prop_left.prop_line == BorderLineKind.DOUBLE
        assert tp.prop_inner_border_table.prop_right.prop_color == StandardColor.BRICK_LIGHT2
        assert tp.prop_inner_border_table.prop_right.prop_line == BorderLineKind.DOTTED
        assert tp.prop_inner_border_table.prop_top.prop_color == StandardColor.BLUE_LIGHT2
        assert tp.prop_inner_border_table.prop_top.prop_line == BorderLineKind.SOLID
        assert tp.prop_inner_border_table.prop_bottom.prop_color == StandardColor.BLUE_LIGHT2
        assert tp.prop_inner_border_table.prop_bottom.prop_line == BorderLineKind.SOLID

        side = Side(line=BorderLineKind.DOUBLE_THIN, color=StandardColor.RED_DARK1, width=LineSize.THIN)
        style = style.fmt_horizontal(side).fmt_vertical(side)
        style.apply(table)

        tp = Borders.from_obj(table)
        assert tp.prop_inner_border_table is not None
        assert tp.prop_inner_shadow is not None
        assert tp.prop_inner_padding is not None
        assert tp.prop_inner_border_table.prop_left.prop_color == StandardColor.GREEN_DARK3
        assert tp.prop_inner_border_table.prop_left.prop_line == BorderLineKind.DOUBLE
        assert tp.prop_inner_border_table.prop_right.prop_color == StandardColor.BRICK_LIGHT2
        assert tp.prop_inner_border_table.prop_right.prop_line == BorderLineKind.DOTTED
        assert tp.prop_inner_border_table.prop_top.prop_color == StandardColor.BLUE_LIGHT2
        assert tp.prop_inner_border_table.prop_top.prop_line == BorderLineKind.SOLID
        assert tp.prop_inner_border_table.prop_bottom.prop_color == StandardColor.BLUE_LIGHT2
        assert tp.prop_inner_border_table.prop_bottom.prop_line == BorderLineKind.SOLID
        assert tp.prop_inner_border_table.prop_horizontal.prop_color == StandardColor.RED_DARK1
        assert tp.prop_inner_border_table.prop_horizontal.prop_line == BorderLineKind.DOUBLE_THIN
        assert tp.prop_inner_border_table.prop_vertical.prop_color == StandardColor.RED_DARK1
        assert tp.prop_inner_border_table.prop_vertical.prop_line == BorderLineKind.DOUBLE_THIN

        style = Borders(merge_adjacent=False)
        style.apply(table)
        tp = Borders.from_obj(table)
        assert tp.prop_merge_adjacent == False

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_write_shadow(loader) -> None:
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

        style = Borders(shadow=Shadow(), merge_adjacent=True)

        table = Write.add_table(cursor=cursor, table_data=tbl_data, first_row_header=False, styles=(style,))

        tp = Borders.from_obj(table)
        assert tp.prop_inner_border_table is not None
        assert tp.prop_inner_shadow is not None
        assert tp.prop_inner_padding is not None
        assert tp.prop_inner_shadow.prop_color == StandardColor.GRAY
        assert tp.prop_inner_shadow.prop_location == ShadowLocation.BOTTOM_RIGHT
        assert tp.prop_merge_adjacent == True

        style = style.fmt_shadow(Shadow(location=ShadowLocation.BOTTOM_RIGHT, color=StandardColor.BLUE_LIGHT3))
        style.apply(table)

        tp = Borders.from_obj(table)
        assert tp.prop_inner_shadow.prop_color == StandardColor.BLUE_LIGHT3
        assert tp.prop_inner_shadow.prop_location == ShadowLocation.BOTTOM_RIGHT

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_write_padding(loader) -> None:
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

        padding100 = UnitMM100.from_mm(3.5)
        style = Borders(padding=Padding(all=padding100), merge_adjacent=True)

        table = Write.add_table(cursor=cursor, table_data=tbl_data, first_row_header=False, styles=(style,))

        tp = Borders.from_obj(table)
        assert tp.prop_inner_border_table is not None
        assert tp.prop_inner_shadow is not None
        assert tp.prop_inner_padding is not None
        assert tp.prop_merge_adjacent == True
        assert tp.prop_inner_padding.prop_left.get_value_mm100() in range(
            padding100.value - 2, padding100.value + 3
        )  # +- 2
        assert tp.prop_inner_padding.prop_right.get_value_mm100() in range(
            padding100.value - 2, padding100.value + 3
        )  # +- 2
        assert tp.prop_inner_padding.prop_top.get_value_mm100() in range(
            padding100.value - 2, padding100.value + 3
        )  # +- 2
        assert tp.prop_inner_padding.prop_bottom.get_value_mm100() in range(
            padding100.value - 2, padding100.value + 3
        )  # +- 2

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
