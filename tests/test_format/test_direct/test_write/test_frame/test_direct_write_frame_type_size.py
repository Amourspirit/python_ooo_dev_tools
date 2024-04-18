from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.frame.type import Size, RelativeKind, RelativeSize, AbsoluteSize
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.color import StandardColor
from ooodev.office.write import Write
from ooodev.units.unit_mm import UnitMM


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

        style = Size(width=AbsoluteSize(200.3), height=AbsoluteSize(50.2), auto_height=True, auto_width=False)

        frame = Write.add_text_frame(
            cursor=cursor, ypos=UnitMM(10.2), text=para_text, width=UnitMM(60), height=UnitMM(40), styles=(style,)
        )

        # style.apply(frame)

        f_style = Size.from_obj(frame)

        assert f_style.prop_height == style.prop_height
        assert f_style.prop_width == style.prop_width
        assert f_style.prop_auto_width == style.prop_auto_width
        assert f_style.prop_auto_height == style.prop_auto_height

        style = Size(
            width=RelativeSize(60, RelativeKind.PAGE),
            height=RelativeSize(40, RelativeKind.PARAGRAPH),
            auto_height=False,
            auto_width=False,
        )

        style.apply(frame)

        f_style = Size.from_obj(frame)
        assert f_style.prop_height == style.prop_height
        assert f_style.prop_width == style.prop_width
        assert f_style.prop_auto_width == style.prop_auto_width
        assert f_style.prop_auto_height == style.prop_auto_height

        style = Size(
            width=AbsoluteSize(50),
            height=RelativeSize(30, RelativeKind.PAGE),
            auto_height=True,
            auto_width=False,
        )

        style.apply(frame)

        f_style = Size.from_obj(frame)
        assert f_style.prop_height == style.prop_height
        assert f_style.prop_width == style.prop_width
        assert f_style.prop_auto_width == style.prop_auto_width
        assert f_style.prop_auto_height == style.prop_auto_height

        style = Size(
            width=RelativeSize(70, RelativeKind.PARAGRAPH),
            height=AbsoluteSize(70.0),
            auto_height=False,
            auto_width=True,
        )

        style.apply(frame)

        f_style = Size.from_obj(frame)
        assert f_style.prop_height == style.prop_height
        assert f_style.prop_width == style.prop_width
        assert f_style.prop_auto_width == style.prop_auto_width
        assert f_style.prop_auto_height == style.prop_auto_height

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
