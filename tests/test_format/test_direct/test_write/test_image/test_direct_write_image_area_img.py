from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.image.area import Img, PresetImageKind, SizeMM
from ooodev.format.writer.direct.image.options import Names
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.images_lo import ImagesLo


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
        style_names = Names(name="skinner", desc="Skinner Pointing", alt="Pointer")
        style = Img.from_preset(preset=PresetImageKind.COLORFUL_PEBBLES)

        _ = Write.add_image_link(
            doc=doc,
            cursor=cursor,
            fnm=im_fnm,
            width=img_size.Width,
            height=img_size.Height,
            styles=(style_names, style),
        )

        graphics = Write.get_graphic_links(doc=doc)
        assert graphics is not None
        assert graphics.hasByName(style_names.prop_name)
        graphic = graphics.getByName(style_names.prop_name)

        f_style = Img.from_obj(graphic)
        point = PresetImageKind.COLORFUL_PEBBLES._get_point()
        assert f_style.prop_is_size_mm
        size = f_style.prop_size
        assert isinstance(size, SizeMM)
        assert round(size.width * 100) in range(point.x - 2, point.x + 3)  # +- 2
        assert round(size.height * 100) in range(point.y - 2, point.y + 3)  # +- 2

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
