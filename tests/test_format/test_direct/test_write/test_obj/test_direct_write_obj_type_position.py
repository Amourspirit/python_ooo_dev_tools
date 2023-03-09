from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.obj.options import Names
from ooodev.format.writer.direct.obj.type import (
    Size,
    RelativeKind,
    RelativeSize,
    AbsoluteSize,
    Position,
    HoriOrient,
    VertOrient,
    RelHoriOrient,
    RelVertOrient,
    Horizontal,
    Vertical,
)
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.images_lo import ImagesLo
from ooodev.utils.data_type.unit_mm100 import UnitMM100
from ooodev.office.write import Write


def test_write(loader, formula_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)

        size_style = Size(
            width=AbsoluteSize(UnitMM100(2000)),
            height=AbsoluteSize(UnitMM100(1000)),
        )

        style_name = Names(name="formula", desc="Just a test Formula", alt="A real Formula")
        style_position = Position(
            horizontal=Horizontal(position=HoriOrient.CENTER, rel=RelHoriOrient.LEFT_PARAGRAPH_BORDER),
            vertical=Vertical(position=VertOrient.CENTER, rel=RelVertOrient.PAGE_TEXT_AREA),
            mirror_even=False,
            keep_boundries=False,
        )

        # write formula can't actually change size but we can read and write the properties.
        content = Write.add_formula(
            cursor=cursor, formula=formula_text, styles=(style_name, size_style, style_position)
        )

        f_style = Position.from_obj(content)

        assert f_style.prop_horizontal.position == style_position.prop_horizontal.position
        assert f_style.prop_horizontal.rel == style_position.prop_horizontal.rel
        assert f_style.prop_vertical.position == style_position.prop_vertical.position
        assert f_style.prop_vertical.rel == style_position.prop_vertical.rel
        assert f_style.prop_mirror_even == style_position.prop_mirror_even
        assert f_style.prop_keep_boundries == style_position.prop_keep_boundries

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
