from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.draw import Draw, DrawDoc
from ooodev.format.draw.direct.position_size.position_size import Position, Size, ShapeBasePointKind


def test_draw_size(loader) -> None:
    doc = DrawDoc(Draw.create_draw_doc())
    try:
        slide = doc.get_slide()
        # left_border = slide.component.BorderLeft
        # top_border = slide.component.BorderTop
        width = 36
        height = 36
        x = round(width / 2)
        y = round(height / 2)

        rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        style = Size(width=50, height=50)
        style_pos = Position(8, 8)
        style.apply(rect.component)
        assert style.prop_width.get_value_mm100() == 5000
        assert style.prop_height.get_value_mm100() == 5000

        f_style = Size.from_obj(rect.component)
        assert f_style is not None
        assert f_style.prop_width.get_value_mm100() == 5000
        assert f_style.prop_height.get_value_mm100() == 5000

        pos = Position.from_obj(rect.component)
        assert pos is not None
        assert pos.prop_pos_x.get_value_mm100() == 800
        assert pos.prop_pos_y.get_value_mm100() == 800

        style2 = style.copy()
        style2.prop_base_point = ShapeBasePointKind.TOP_CENTER
        style2.apply(rect.component)

        # no change expected
        pos = Position.from_obj(rect.component)
        assert pos is not None
        assert pos.prop_pos_x.get_value_mm100() == 800
        assert pos.prop_pos_y.get_value_mm100() == 800

        style2.prop_width = 100
        style2.prop_height = 100
        style2.apply(rect.component)

        f_style = Size.from_obj(rect.component)
        assert f_style is not None
        assert f_style.prop_width.get_value_mm100() == 10000
        assert f_style.prop_height.get_value_mm100() == 10000

        pos = Position.from_obj(rect.component)
        assert pos is not None
        assert pos.prop_pos_x.get_value_mm100() == -1700
        assert pos.prop_pos_y.get_value_mm100() == 800

        # reset
        style.apply(rect.component)
        style_pos.apply(rect.component)

        style2.prop_base_point = ShapeBasePointKind.TOP_RIGHT
        style2.apply(rect.component)

        f_style = Size.from_obj(rect.component)
        assert f_style is not None
        assert f_style.prop_width.get_value_mm100() == 10000
        assert f_style.prop_height.get_value_mm100() == 10000

        pos = Position.from_obj(rect.component)
        assert pos is not None
        assert pos.prop_pos_x.get_value_mm100() == -4200
        assert pos.prop_pos_y.get_value_mm100() == 800

        # reset
        style.apply(rect.component)
        style_pos.apply(rect.component)

        style2.prop_base_point = ShapeBasePointKind.CENTER_LEFT
        style2.apply(rect.component)

        f_style = Size.from_obj(rect.component)
        assert f_style is not None
        assert f_style.prop_width.get_value_mm100() == 10000
        assert f_style.prop_height.get_value_mm100() == 10000

        pos = Position.from_obj(rect.component)
        assert pos is not None
        assert pos.prop_pos_x.get_value_mm100() == 800
        assert pos.prop_pos_y.get_value_mm100() == -1700

        # reset
        style.apply(rect.component)
        style_pos.apply(rect.component)

        style2.prop_base_point = ShapeBasePointKind.CENTER
        style2.apply(rect.component)

        f_style = Size.from_obj(rect.component)
        assert f_style is not None
        assert f_style.prop_width.get_value_mm100() == 10000
        assert f_style.prop_height.get_value_mm100() == 10000

        pos = Position.from_obj(rect.component)
        assert pos is not None
        assert pos.prop_pos_x.get_value_mm100() == -1700
        assert pos.prop_pos_y.get_value_mm100() == -1700

        # reset
        style.apply(rect.component)
        style_pos.apply(rect.component)

        style2.prop_base_point = ShapeBasePointKind.CENTER_RIGHT
        style2.apply(rect.component)

        f_style = Size.from_obj(rect.component)
        assert f_style is not None
        assert f_style.prop_width.get_value_mm100() == 10000
        assert f_style.prop_height.get_value_mm100() == 10000

        pos = Position.from_obj(rect.component)
        assert pos is not None
        assert pos.prop_pos_x.get_value_mm100() == -4200
        assert pos.prop_pos_y.get_value_mm100() == -1700

        # reset
        style.apply(rect.component)
        style_pos.apply(rect.component)

        style2.prop_base_point = ShapeBasePointKind.BOTTOM_LEFT
        style2.apply(rect.component)

        f_style = Size.from_obj(rect.component)
        assert f_style is not None
        assert f_style.prop_width.get_value_mm100() == 10000
        assert f_style.prop_height.get_value_mm100() == 10000

        pos = Position.from_obj(rect.component)
        assert pos is not None
        assert pos.prop_pos_x.get_value_mm100() == 800
        assert pos.prop_pos_y.get_value_mm100() == -4200

        # reset
        style.apply(rect.component)
        style_pos.apply(rect.component)

        style2.prop_base_point = ShapeBasePointKind.BOTTOM_CENTER
        style2.apply(rect.component)

        f_style = Size.from_obj(rect.component)
        assert f_style is not None
        assert f_style.prop_width.get_value_mm100() == 10000
        assert f_style.prop_height.get_value_mm100() == 10000

        pos = Position.from_obj(rect.component)
        assert pos is not None
        assert pos.prop_pos_x.get_value_mm100() == -1700
        assert pos.prop_pos_y.get_value_mm100() == -4200

        # reset
        style.apply(rect.component)
        style_pos.apply(rect.component)

        style2.prop_base_point = ShapeBasePointKind.BOTTOM_RIGHT
        style2.apply(rect.component)

        f_style = Size.from_obj(rect.component)
        assert f_style is not None
        assert f_style.prop_width.get_value_mm100() == 10000
        assert f_style.prop_height.get_value_mm100() == 10000

        pos = Position.from_obj(rect.component)
        assert pos is not None
        assert pos.prop_pos_x.get_value_mm100() == -4200
        assert pos.prop_pos_y.get_value_mm100() == -4200
    finally:
        doc.close_doc()
