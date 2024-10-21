from __future__ import annotations
from typing import cast
from pathlib import Path
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.image.options import Names
from ooodev.format.writer.direct.image.type import (
    Anchor,
    AnchorKind,
    Size,
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
from ooodev.utils.images_lo import ImagesLo
from ooodev.units.unit_mm100 import UnitMM100
from ooodev.office.write import Write


def test_write(loader, fix_image_path) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        im_fnm = cast(Path, fix_image_path("skinner.png"))
        cursor = Write.get_cursor(doc)
        img_size = ImagesLo.get_size_100mm(im_fnm=im_fnm)
        size_style = Size(
            width=AbsoluteSize(UnitMM100(img_size.Width)),
            height=AbsoluteSize(UnitMM100(img_size.Height)),
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
        style_name = Names(name="skinner", desc="Skinner Pointing", alt="Pointer")

        _ = Write.add_image_link(
            doc=doc,
            cursor=cursor,
            fnm=im_fnm,
            styles=(
                style_name,
                style_anchor,
                size_style,
                style_position,
            ),
        )

        graphics = Write.get_graphic_links(doc=doc)
        assert graphics is not None
        assert graphics.hasByName(style_name.prop_name)
        graphic = graphics.getByName(style_name.prop_name)

        f_style_anchor = Anchor.from_obj(graphic)
        assert f_style_anchor.prop_anchor == style_anchor.prop_anchor

        f_style_position = Position.from_obj(graphic)

        assert f_style_position.prop_horizontal.position == style_position.prop_horizontal.position
        assert f_style_position.prop_horizontal.rel == style_position.prop_horizontal.rel
        assert f_style_position.prop_vertical.position == style_position.prop_vertical.position
        assert f_style_position.prop_vertical.rel == style_position.prop_vertical.rel
        assert f_style_position.prop_mirror_even == style_position.prop_mirror_even
        assert f_style_position.prop_keep_boundaries == style_position.prop_keep_boundaries

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
