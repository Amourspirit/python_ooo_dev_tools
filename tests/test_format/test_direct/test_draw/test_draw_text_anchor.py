from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.draw import Draw, DrawDoc
from ooodev.format.draw.direct.text.text import TextAnchor, ShapeBasePointKind


def test_draw_shape_text_anchor(loader) -> None:
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
        txt_anchor = TextAnchor(anchor_point=ShapeBasePointKind.TOP_LEFT)
        txt_anchor.apply(rect.component)
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is False
        assert f_style.prop_anchor_point == ShapeBasePointKind.TOP_LEFT

        txt_anchor = TextAnchor(anchor_point=ShapeBasePointKind.TOP_CENTER)
        txt_anchor.apply(rect.component)
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is False
        assert f_style.prop_anchor_point == ShapeBasePointKind.TOP_CENTER

        txt_anchor.prop_full_width = True
        txt_anchor.apply(rect.component)
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is True
        assert f_style.prop_anchor_point == ShapeBasePointKind.TOP_CENTER

        txt_anchor = TextAnchor(anchor_point=ShapeBasePointKind.TOP_RIGHT)
        txt_anchor.apply(rect.component)
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is False
        assert f_style.prop_anchor_point == ShapeBasePointKind.TOP_RIGHT

        txt_anchor = TextAnchor(anchor_point=ShapeBasePointKind.CENTER_LEFT)
        txt_anchor.apply(rect.component)
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is False
        assert f_style.prop_anchor_point == ShapeBasePointKind.CENTER_LEFT

        txt_anchor = TextAnchor(anchor_point=ShapeBasePointKind.CENTER, full_width=False)
        txt_anchor.apply(rect.component)
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is False
        assert f_style.prop_anchor_point == ShapeBasePointKind.CENTER

        txt_anchor = TextAnchor(anchor_point=ShapeBasePointKind.CENTER, full_width=True)
        txt_anchor.apply(rect.component)
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is True
        assert f_style.prop_anchor_point == ShapeBasePointKind.CENTER

        txt_anchor = TextAnchor(anchor_point=ShapeBasePointKind.CENTER_RIGHT)
        txt_anchor.apply(rect.component)
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is False
        assert f_style.prop_anchor_point == ShapeBasePointKind.CENTER_RIGHT

        txt_anchor = TextAnchor(anchor_point=ShapeBasePointKind.BOTTOM_LEFT)
        txt_anchor.apply(rect.component)
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is False
        assert f_style.prop_anchor_point == ShapeBasePointKind.BOTTOM_LEFT

        txt_anchor = TextAnchor(anchor_point=ShapeBasePointKind.BOTTOM_CENTER)
        txt_anchor.apply(rect.component)
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is False
        assert f_style.prop_anchor_point == ShapeBasePointKind.BOTTOM_CENTER

        txt_anchor = TextAnchor(anchor_point=ShapeBasePointKind.BOTTOM_CENTER, full_width=True)
        txt_anchor.apply(rect.component)
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is True
        assert f_style.prop_anchor_point == ShapeBasePointKind.BOTTOM_CENTER

        txt_anchor = TextAnchor(anchor_point=ShapeBasePointKind.BOTTOM_RIGHT)
        txt_anchor.apply(rect.component)
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is False
        assert f_style.prop_anchor_point == ShapeBasePointKind.BOTTOM_RIGHT

    finally:
        doc.close_doc()
