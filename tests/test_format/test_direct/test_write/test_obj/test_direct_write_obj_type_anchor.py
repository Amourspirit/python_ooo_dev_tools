from __future__ import annotations

import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.obj.options import Names
from ooodev.format.writer.direct.obj.type import (
    Anchor,
    AnchorKind,
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
from ooodev.utils.data_type.unit_inch import UnitInch
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
            width=AbsoluteSize(UnitInch(1.0)),
            height=AbsoluteSize(UnitInch(0.5)),
        )
        style_position = Position(
            horizontal=Horizontal(
                position=HoriOrient.FROM_LEFT_OR_INSIDE, rel=RelHoriOrient.PARAGRAPH_AREA, amount=10.0
            ),
            vertical=Vertical(position=VertOrient.FROM_TOP_OR_BOTTOM, rel=RelVertOrient.MARGIN, amount=12.0),
            mirror_even=False,
            keep_boundries=False,
        )
        style_anchor = Anchor(anchor=AnchorKind.AT_CHARACTER)
        style_name = Names(name="formula", desc="Just a test Formula", alt="A real Formula")

        # writer formula can't actually change size but we can read and write the properties.
        content = Write.add_formula(
            cursor=cursor, formula=formula_text, styles=(style_name, size_style, style_anchor, style_position)
        )

        f_style_anchor = Anchor.from_obj(content)
        assert f_style_anchor.prop_anchor == style_anchor.prop_anchor

        f_style_position = Position.from_obj(content)

        assert f_style_position.prop_horizontal.position == style_position.prop_horizontal.position
        assert f_style_position.prop_horizontal.rel == style_position.prop_horizontal.rel
        assert f_style_position.prop_vertical.position == style_position.prop_vertical.position
        assert f_style_position.prop_vertical.rel == style_position.prop_vertical.rel
        assert f_style_position.prop_mirror_even == style_position.prop_mirror_even
        assert f_style_position.prop_keep_boundries == style_position.prop_keep_boundries

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
