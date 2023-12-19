from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.draw import Draw, DrawDoc
from ooodev.format.draw.direct.position_size.position_size import Position, ShapeBasePointKind


def test_draw_position(loader) -> None:
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
        pos = Position(pos_x=8, pos_y=8)
        pos.apply(rect.component)
        assert pos.prop_pos_x.get_value_mm100() in [800 - 2 + i for i in range(5)]  # plus or minus 2

        pos2 = Position.from_obj(rect.component)
        assert pos2 is not None
        assert pos2.prop_pos_x.get_value_mm100() in [800 - 2 + i for i in range(5)]

        pos = Position(pos_x=8, pos_y=8, base_point=ShapeBasePointKind.TOP_CENTER)
        pos.apply(rect.component)

        pos2 = Position.from_obj(rect.component)
        assert pos2 is not None
        assert pos2.prop_pos_x.get_value_mm100() in [2600 - 2 + i for i in range(5)]
        assert pos2.prop_pos_y.get_value_mm100() in [800 - 2 + i for i in range(5)]

        pos = Position(pos_x=8, pos_y=8, base_point=ShapeBasePointKind.TOP_RIGHT)
        pos.apply(rect.component)

        pos2 = Position.from_obj(rect.component)
        assert pos2 is not None
        assert pos2.prop_pos_x.get_value_mm100() in [4400 - 2 + i for i in range(5)]
        assert pos2.prop_pos_y.get_value_mm100() in [800 - 2 + i for i in range(5)]

        pos = Position(pos_x=8, pos_y=8, base_point=ShapeBasePointKind.CENTER_LEFT)
        pos.apply(rect.component)

        pos2 = Position.from_obj(rect.component)
        assert pos2 is not None
        assert pos2.prop_pos_x.get_value_mm100() in [800 - 2 + i for i in range(5)]
        assert pos2.prop_pos_y.get_value_mm100() in [2600 - 2 + i for i in range(5)]

        pos = Position(pos_x=8, pos_y=8, base_point=ShapeBasePointKind.CENTER)
        pos.apply(rect.component)

        pos2 = Position.from_obj(rect.component)
        assert pos2 is not None
        assert pos2.prop_pos_x.get_value_mm100() in [2600 - 2 + i for i in range(5)]
        assert pos2.prop_pos_y.get_value_mm100() in [2600 - 2 + i for i in range(5)]

        pos = Position(pos_x=8, pos_y=8, base_point=ShapeBasePointKind.CENTER_RIGHT)
        pos.apply(rect.component)

        pos2 = Position.from_obj(rect.component)
        assert pos2 is not None
        assert pos2.prop_pos_x.get_value_mm100() in [4400 - 2 + i for i in range(5)]
        assert pos2.prop_pos_y.get_value_mm100() in [2600 - 2 + i for i in range(5)]

        pos = Position(pos_x=8, pos_y=8, base_point=ShapeBasePointKind.BOTTOM_LEFT)
        pos.apply(rect.component)

        pos2 = Position.from_obj(rect.component)
        assert pos2 is not None
        assert pos2.prop_pos_x.get_value_mm100() in [800 - 2 + i for i in range(5)]
        assert pos2.prop_pos_y.get_value_mm100() in [4400 - 2 + i for i in range(5)]

        pos = Position(pos_x=8, pos_y=8, base_point=ShapeBasePointKind.BOTTOM_CENTER)
        pos.apply(rect.component)

        pos2 = Position.from_obj(rect.component)
        assert pos2 is not None
        assert pos2.prop_pos_x.get_value_mm100() in [2600 - 2 + i for i in range(5)]
        assert pos2.prop_pos_y.get_value_mm100() in [4400 - 2 + i for i in range(5)]

        pos = Position(pos_x=8, pos_y=8, base_point=ShapeBasePointKind.BOTTOM_RIGHT)
        pos.apply(rect.component)

        pos2 = Position.from_obj(rect.component)
        assert pos2 is not None
        assert pos2.prop_pos_x.get_value_mm100() in [4400 - 2 + i for i in range(5)]
        assert pos2.prop_pos_y.get_value_mm100() in [4400 - 2 + i for i in range(5)]

    finally:
        doc.close_doc()
