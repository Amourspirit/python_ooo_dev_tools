from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.units.unit_convert import UnitConvert
from ooodev.format.writer.direct.frame.type import (
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
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.color import StandardColor
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.units.unit_mm import UnitMM
from ooodev.office.write import Write


def test_write(loader, para_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)
        if not Lo.bridge_connector.headless:
            Write.append_para(cursor=cursor, text=para_text)

        text_width = Write.get_page_text_width(doc)
        size_style = Size(
            width=AbsoluteSize(UnitMM100(text_width)), height=AbsoluteSize(50.2), auto_height=True, auto_width=False
        )
        style_position = Position(
            horizontal=Horizontal(
                position=HoriOrient.FROM_LEFT_OR_INSIDE, rel=RelHoriOrient.PARAGRAPH_AREA, amount=10.0
            ),
            vertical=Vertical(position=VertOrient.FROM_TOP_OR_BOTTOM, rel=RelVertOrient.MARGIN, amount=12.0),
            mirror_even=False,
            keep_boundaries=False,
        )
        style_anchor = Anchor(anchor=AnchorKind.AT_CHARACTER)

        frame = Write.add_text_frame(
            cursor=cursor,
            ypos=UnitMM(10),
            text=para_text,
            width=UnitMM(60),
            height=UnitMM(40),
            styles=(style_anchor, size_style, style_position),
        )

        f_style_anchor = Anchor.from_obj(frame)
        assert f_style_anchor.prop_anchor == style_anchor.prop_anchor

        f_style_position = Position.from_obj(frame)

        assert f_style_position.prop_horizontal.position == style_position.prop_horizontal.position
        assert f_style_position.prop_horizontal.rel == style_position.prop_horizontal.rel
        assert f_style_position.prop_vertical.position == style_position.prop_vertical.position
        assert f_style_position.prop_vertical.rel == style_position.prop_vertical.rel
        assert f_style_position.prop_mirror_even == style_position.prop_mirror_even
        assert f_style_position.prop_keep_boundaries == style_position.prop_keep_boundaries

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
