from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.frame.type import (
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
from ooodev.office.write import Write
from ooodev.utils.data_type.unit_mm import UnitMM


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

        size_style = Size(width=AbsoluteSize(200.3), height=AbsoluteSize(50.2), auto_height=True, auto_width=False)
        style = Position(
            horizontal=Horizontal(position=HoriOrient.CENTER, rel=RelHoriOrient.LEFT_PARAGRAPH_BORDER),
            vertical=Vertical(position=VertOrient.CENTER, rel=RelVertOrient.PAGE_TEXT_AREA),
            mirror_even=False,
            keep_boundaries=False,
        )

        frame = Write.add_text_frame(
            cursor=cursor,
            ypos=UnitMM(10.2),
            text=para_text,
            width=UnitMM(60),
            height=UnitMM(40),
            styles=(size_style, style),
        )

        # style.apply(frame)

        f_style = Position.from_obj(frame)

        assert f_style.prop_horizontal.position == style.prop_horizontal.position
        assert f_style.prop_horizontal.rel == style.prop_horizontal.rel
        assert f_style.prop_vertical.position == style.prop_vertical.position
        assert f_style.prop_vertical.rel == style.prop_vertical.rel
        assert f_style.prop_mirror_even == style.prop_mirror_even
        assert f_style.prop_keep_boundaries == style.prop_keep_boundaries

        style = Position(
            horizontal=Horizontal(position=HoriOrient.LEFT_OR_INSIDE, rel=RelHoriOrient.ENTIRE_PAGE),
            vertical=Vertical(position=VertOrient.BOTTOM, rel=RelVertOrient.PAGE_TEXT_AREA),
            mirror_even=True,
            keep_boundaries=True,
        )

        style.apply(frame)
        f_style = Position.from_obj(frame)
        assert f_style.prop_horizontal.position == style.prop_horizontal.position
        assert f_style.prop_horizontal.rel == style.prop_horizontal.rel
        assert f_style.prop_vertical.position == style.prop_vertical.position
        assert f_style.prop_vertical.rel == style.prop_vertical.rel
        assert f_style.prop_mirror_even == style.prop_mirror_even
        assert f_style.prop_keep_boundaries == style.prop_keep_boundaries

        style = Position(
            horizontal=Horizontal(position=HoriOrient.FROM_LEFT_OR_INSIDE, rel=RelHoriOrient.ENTIRE_PAGE),
            vertical=Vertical(position=VertOrient.CENTER, rel=RelVertOrient.ENTIRE_PAGE_OR_ROW),
            mirror_even=False,
            keep_boundaries=False,
        )

        style.apply(frame)
        f_style = Position.from_obj(frame)
        assert f_style.prop_horizontal.position == style.prop_horizontal.position
        assert f_style.prop_horizontal.rel == style.prop_horizontal.rel
        assert f_style.prop_vertical.position == style.prop_vertical.position
        assert f_style.prop_vertical.rel == style.prop_vertical.rel
        assert f_style.prop_mirror_even == style.prop_mirror_even
        assert f_style.prop_keep_boundaries == style.prop_keep_boundaries

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
