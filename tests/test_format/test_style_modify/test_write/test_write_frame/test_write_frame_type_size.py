from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.modify.frame.type import Size, RelativeKind, RelativeSize, AbsoluteSize
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
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

        style = Size(width=AbsoluteSize(12.3), height=AbsoluteSize(10.2), auto_height=True, auto_width=False)

        style.apply(doc)
        # props = style.get_style_props(doc)

        f_style = Size.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_height == style.prop_inner.prop_height
        assert f_style.prop_inner.prop_width == style.prop_inner.prop_width
        assert f_style.prop_inner.prop_auto_width == style.prop_inner.prop_auto_width
        assert f_style.prop_inner.prop_auto_height == style.prop_inner.prop_auto_height

        style = Size(
            width=RelativeSize(10, RelativeKind.PAGE),
            height=RelativeSize(5, RelativeKind.PARAGRAPH),
            auto_height=False,
            auto_width=False,
        )

        style.apply(doc)

        f_style = Size.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_height == style.prop_inner.prop_height
        assert f_style.prop_inner.prop_width == style.prop_inner.prop_width
        assert f_style.prop_inner.prop_auto_width == style.prop_inner.prop_auto_width
        assert f_style.prop_inner.prop_auto_height == style.prop_inner.prop_auto_height

        style = Size(
            width=AbsoluteSize(11.3),
            height=RelativeSize(7, RelativeKind.PAGE),
            auto_height=True,
            auto_width=False,
        )

        style.apply(doc)

        f_style = Size.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_height == style.prop_inner.prop_height
        assert f_style.prop_inner.prop_width == style.prop_inner.prop_width
        assert f_style.prop_inner.prop_auto_width == style.prop_inner.prop_auto_width
        assert f_style.prop_inner.prop_auto_height == style.prop_inner.prop_auto_height

        style = Size(
            width=RelativeSize(7, RelativeKind.PARAGRAPH),
            height=AbsoluteSize(11.3),
            auto_height=False,
            auto_width=True,
        )

        style.apply(doc)

        f_style = Size.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_height == style.prop_inner.prop_height
        assert f_style.prop_inner.prop_width == style.prop_inner.prop_width
        assert f_style.prop_inner.prop_auto_width == style.prop_inner.prop_auto_width
        assert f_style.prop_inner.prop_auto_height == style.prop_inner.prop_auto_height

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
