import os
import pytest
from pathlib import Path
from typing import cast, TYPE_CHECKING


# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

def test_get_bitmap(loader, fix_image_path) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.write import Write
    from ooodev.utils.images_lo import ImagesLo
    
    im_fnm = cast(Path, fix_image_path("skinner.png"))
    # needs an office document to use get_bitmap
    doc = Write.create_doc(loader)
    try:
        bitmap = ImagesLo.get_bitmap(im_fnm)
        b_size = bitmap.getSize()
        assert b_size.Height == 274
        assert b_size.Width == 319

        img_size = ImagesLo.get_size_100mm(im_fnm=im_fnm)
        assert img_size.Height == 5751
        assert img_size.Width == 6092
    finally:
        Lo.close_doc(doc, False)

def test_load_graphic(loader, fix_image_path) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.write import Write
    from ooodev.utils.images_lo import ImagesLo
    from ooodev.utils.props import Props
    from ooodev.utils.info import Info
    from functools import partial
    if TYPE_CHECKING:
        from com.sun.star.graphic import Graphic
    im_fnm = cast(Path, fix_image_path("skinner.png"))
    doc = Write.create_doc(loader)
    try:
        graphic = ImagesLo.load_graphic_file(im_fnm)
        assert graphic is not None
        get_prop = partial(Props.get_property, graphic)
        size = get_prop("SizePixel")
        assert size.Height == 274
        assert size.Width == 319
        alpha = get_prop("Alpha")
        assert alpha
        bits_per_pixel = get_prop("BitsPerPixel")
        assert bits_per_pixel == 24
        mime_type = get_prop("MimeType")
        assert mime_type == "image/png"

        assert Info.support_service(graphic, "com.sun.star.graphic.Graphic")
        g = cast("Graphic", graphic)

        assert g.Alpha == True
        assert g.BitsPerPixel == 24
        assert g.MimeType == "image/png"
        size_100 = g.Size100thMM
        assert size_100.Height == 5751
        assert size_100.Width == 6092
        assert g.Transparent
    finally:
        Lo.close_doc(doc, False)

def test_get_size_pixels(loader, fix_image_path) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.write import Write
    from ooodev.utils.images_lo import ImagesLo
    
    im_fnm = cast(Path, fix_image_path("skinner.png"))
    # needs an office document to use get_bitmap
    doc = Write.create_doc(loader)
    try:
        size = ImagesLo.get_size_pixels(im_fnm)
        assert size.Height == 274
        assert size.Width == 319
    finally:
        Lo.close_doc(doc, False)

def _test_add_image_link(loader, fix_image_path) -> None:
    from ooodev.utils.lo import Lo
    from ooodev.office.write import Write
    from ooodev.utils.images_lo import ImagesLo
    from ooodev.utils.gui import GUI
    visible = True
    delay = 2000
    im_fnm = cast(Path, fix_image_path("skinner.png"))
    # needs an office document to use get_bitmap
    doc = Write.create_doc(loader)
    try:
        if visible:
            GUI.set_visible(visible, doc)
        cursor = Write.get_cursor(doc)
        Write.add_image_link(doc, cursor, im_fnm)
        graphics = Write.get_text_graphics(doc)
        assert graphics is None
        Lo.delay(delay)
        
        Write.append(cursor, "Image as a shape: ")
        Write.add_image_shape(doc=doc, cursor=cursor, fnm=im_fnm)
        graphics = Write.get_text_graphics(doc)
        assert graphics is None
        Lo.delay(delay)
        
        graphic = ImagesLo.load_graphic_file(im_fnm)
        gl = ImagesLo.load_graphic_link(graphic)
        assert gl is not None
    finally:
        Lo.close_doc(doc, False)