from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.image.wrap import (
    Options,
    Settings,
    WrapTextMode,
)
from ooodev.format.writer.direct.image.options import Names
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.images_lo import ImagesLo


def test_write(loader, fix_image_path) -> None:
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
        style_settings = Settings(mode=WrapTextMode.THROUGH)
        # background can only be set to true when mode is Through
        style = Options(first=True, background=True)

        _ = Write.add_image_link(
            doc=doc,
            cursor=cursor,
            fnm=im_fnm,
            width=img_size.Width,
            height=img_size.Height,
            styles=(style_names, style_settings, style),
        )

        graphics = Write.get_graphic_links(doc=doc)
        assert graphics is not None
        assert graphics.hasByName(style_names.prop_name)
        graphic = graphics.getByName(style_names.prop_name)

        f_style = Options.from_obj(graphic)
        assert f_style.prop_first
        assert f_style.prop_first == style.prop_first
        assert f_style.prop_background
        assert f_style.prop_background == style.prop_background

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
