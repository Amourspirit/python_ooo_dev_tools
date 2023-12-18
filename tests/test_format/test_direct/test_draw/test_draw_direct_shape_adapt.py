from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.draw import Draw, DrawDoc
from ooodev.format.draw.direct.position_size.position_size import Adapt


def test_draw_adapt(loader) -> None:
    doc = DrawDoc(Draw.create_draw_doc())
    try:
        slide = doc.get_slide()
        width = 36
        height = 36
        x = round(width / 2)
        y = round(height / 2)

        text = slide.draw_text(msg="Hello World", x=x, y=y, width=width, height=height)
        style = Adapt(fit_height=True, fit_width=True)
        style.apply(text.component)
        assert style.prop_fit_height
        assert style.prop_fit_height

        f_style = Adapt.from_obj(text.component)
        assert f_style.prop_fit_height
        assert f_style.prop_fit_height

        style.prop_fit_height = False
        style.prop_fit_width = False
        style.apply(text.component)

        f_style = Adapt.from_obj(text.component)
        assert f_style.prop_fit_height is False
        assert f_style.prop_fit_height is False

        # does nothing because it is not supported, only text boxes.
        # rect = slide.draw_rectangle(x=x + 10, y=y + 10, width=width, height=height)
        # style.apply(rect.component)

    finally:
        doc.close_doc()
