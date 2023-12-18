from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.draw import Draw, DrawDoc
from ooodev.format.draw.direct.position_size.position_size import Protect


def test_draw_protect(loader) -> None:
    doc = DrawDoc(Draw.create_draw_doc())
    try:
        slide = doc.get_slide()
        width = 36
        height = 36
        x = round(width / 2)
        y = round(height / 2)

        rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        style = Protect(size=True)
        style2 = Protect(position=False, size=False)
        style.apply(rect.component)
        assert style.prop_size
        assert style.prop_position is None

        f_style = Protect.from_obj(rect.component)
        assert f_style is not None
        assert f_style.prop_size
        assert f_style.prop_position is False

        # reset
        style2.apply(rect.component)

        f_style = Protect.from_obj(rect.component)
        assert f_style.prop_size is False
        assert f_style.prop_position is False

        style.prop_size = False
        style.prop_position = True  # sets size to True
        assert style.prop_size is True
        assert style.prop_position is True
        style.prop_size = False  # should be ignored when position is True
        assert style.prop_size is True
        style.apply(rect.component)

        f_style = Protect.from_obj(rect.component)
        assert f_style.prop_size is True
        assert f_style.prop_position is True

    finally:
        doc.close_doc()
