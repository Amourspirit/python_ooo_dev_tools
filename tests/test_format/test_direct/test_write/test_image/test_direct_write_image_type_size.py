from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.image.type import Size, AbsoluteSize
from ooodev.format.writer.direct.image.options import Names
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
        style_size = Size(
            width=AbsoluteSize(UnitMM100(img_size.Width)), height=AbsoluteSize(UnitMM100(img_size.Height))
        )
        style_name = Names(name="skinner", desc="Skinner Pointing", alt="Pointer")

        _ = Write.add_image_link(
            doc=doc,
            cursor=cursor,
            fnm=im_fnm,
            styles=(
                style_name,
                style_size,
            ),
        )

        graphics = Write.get_graphic_links(doc=doc)
        assert graphics is not None
        assert graphics.hasByName(style_name.prop_name)
        graphic = graphics.getByName(style_name.prop_name)

        f_style = Size.from_obj(graphic)

        assert f_style.prop_height == style_size.prop_height
        assert f_style.prop_width == style_size.prop_width

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
