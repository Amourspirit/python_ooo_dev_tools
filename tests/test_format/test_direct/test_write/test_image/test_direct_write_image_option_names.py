from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])


# from com.sun.star.text import XTextFrame
# from com.sun.star.drawing import XShape

from ooodev.format.writer.direct.image.options import Names
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.images_lo import ImagesLo
from ooodev.office.write import Write


def test_write(loader, fix_image_path) -> None:
    # testing prev and next is not currently possible.
    # see Names class and read comments for details.

    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        im_fnm = cast(Path, fix_image_path("skinner.png"))
        cursor = Write.get_cursor(doc)

        style = Names(name="skinner", desc="Skinner Pointing", alt="Pointer")

        img_size = ImagesLo.get_size_100mm(im_fnm=im_fnm)
        _ = Write.add_image_link(
            doc=doc, cursor=cursor, fnm=im_fnm, width=img_size.Width, height=img_size.Height, styles=(style,)
        )

        graphics = Write.get_graphic_links(doc=doc)
        assert graphics is not None
        assert graphics.hasByName(style.prop_name)
        graphic = graphics.getByName(style.prop_name)

        f_style = Names.from_obj(graphic)
        assert f_style.prop_name == style.prop_name
        assert f_style.prop_desc == style.prop_desc
        assert f_style.prop_alt == style.prop_alt

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
