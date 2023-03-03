from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.image.options import Names
from ooodev.format.writer.direct.image.crop import ImageCrop, CropOpt, Size, SizeMM
from ooodev.utils.data_type.unit_mm import UnitMM
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.images_lo import ImagesLo
from ooodev.office.write import Write


def test_write_crop(loader, fix_image_path) -> None:
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

        style_names = Names(name="skinner", desc="Skinner Pointing", alt="Pointer")

        img_size = ImagesLo.get_size_100mm(im_fnm=im_fnm)

        style = ImageCrop(
            crop=CropOpt(all=1.5, keep_scale=False),
            img_size=SizeMM.from_mm100(img_size.width, img_size.height),
        )

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

        f_style = ImageCrop.from_obj(graphic)
        assert f_style.prop_crop_opt == style.prop_crop_opt
        assert f_style.prop_img_size == style.prop_img_size

        sz = ImageCrop.get_image_original_size(graphic)
        sz_mm = SizeMM.from_size_mm100(sz)
        assert f_style.prop_img_size == sz_mm

        style.prop_crop_opt = CropOpt(all=0)
        style.apply(graphic)

        f_style = ImageCrop.from_obj(graphic)
        f_struct = f_style.prop_crop_opt.get_uno_struct()
        struct = style.prop_crop_opt.get_uno_struct()
        # should all be 0.0 values
        assert struct.Bottom == 0.0
        assert f_struct.Bottom == struct.Bottom
        assert f_struct.Left == struct.Left
        assert f_struct.Right == struct.Right
        assert f_struct.Top == struct.Top

        assert f_style.prop_img_size == style.prop_img_size

        style.prop_crop_opt = CropOpt(all=0, keep_scale=True)
        style.apply(graphic)

        f_style = ImageCrop.from_obj(graphic)
        f_struct = f_style.prop_crop_opt.get_uno_struct()
        # should all be 0.0 values
        assert struct.Bottom == 0.0
        assert f_struct.Bottom == struct.Bottom
        assert f_struct.Left == struct.Left
        assert f_struct.Right == struct.Right
        assert f_struct.Top == struct.Top
        assert f_style.prop_img_size == style.prop_img_size

        style = ImageCrop(crop=CropOpt(all=1.5, keep_scale=True))
        style.apply(graphic)

        f_style = ImageCrop.from_obj(graphic)
        assert f_style.prop_crop_opt == style.prop_crop_opt

        style.prop_crop_opt = CropOpt(all=0, keep_scale=False)
        style.apply(graphic)

        f_style = ImageCrop.from_obj(graphic)
        f_struct = f_style.prop_crop_opt.get_uno_struct()
        struct = style.prop_crop_opt.get_uno_struct()
        # should all be 0.0 values
        assert struct.Bottom == 0.0
        assert f_struct.Bottom == struct.Bottom
        assert f_struct.Left == struct.Left
        assert f_struct.Right == struct.Right
        assert f_struct.Top == struct.Top

        # test keep scale
        style = ImageCrop(
            crop=CropOpt(all=1.5, keep_scale=True),
            img_scale=Size(111, 114),
        )
        style.apply(graphic)
        f_style = ImageCrop.from_obj(graphic)

        w_factor = style.prop_img_scale.width / 100
        h_factor = style.prop_img_scale.height / 100
        crop_w = UnitMM(style.prop_crop_opt.prop_left + style.prop_crop_opt.prop_right).get_value_mm100()
        crop_h = UnitMM(style.prop_crop_opt.prop_top + style.prop_crop_opt.prop_bottom).get_value_mm100()
        width = round((img_size.width - crop_w) * w_factor)  # 1/100th mm
        height = round((img_size.height - crop_h) * h_factor)

        f_sz = f_style.prop_img_size.get_size_mm100()
        assert f_sz.width in range(width - 2, width + 5)  # +- 2
        assert f_sz.height in range(height - 2, height + 5)  # +- 2

        style = ImageCrop(
            crop=CropOpt(all=1.5, keep_scale=False),
            img_size=SizeMM.from_mm100(img_size.width + 100, img_size.height + 100),
        )
        f_style = ImageCrop.from_obj(graphic)
        f_struct = f_style.prop_crop_opt.get_uno_struct()
        struct = style.prop_crop_opt.get_uno_struct()
        assert f_struct.Bottom == struct.Bottom
        assert f_struct.Left == struct.Left
        assert f_struct.Right == struct.Right
        assert f_struct.Top == struct.Top

        ImageCrop.reset_image_original_size(graphic)
        f_style = ImageCrop.from_obj(graphic)
        f_struct = f_style.prop_crop_opt.get_uno_struct()
        struct = style.prop_crop_opt.get_uno_struct()
        # should all be 0.0 values
        assert f_struct.Left == 0.0
        assert f_struct.Right == 0.0
        assert f_struct.Top == 0.0
        assert f_struct.Bottom == 0.0
        sz = ImageCrop.get_image_original_size(graphic)
        f_sz = f_style.prop_img_size.get_size_mm100()
        assert f_sz.width in range(sz.width - 2, sz.width + 3)  # +- 2
        assert f_sz.height in range(sz.height - 2, sz.height + 3)  # +- 2

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_write_no_crop(loader, fix_image_path) -> None:
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

        style_names = Names(name="skinner", desc="Skinner Pointing", alt="Pointer")

        img_size = ImagesLo.get_size_100mm(im_fnm=im_fnm)
        img_size_mm = SizeMM.from_mm100(img_size.width, img_size.height)

        # test image size only
        style = ImageCrop(img_size=SizeMM.from_mm100(img_size.width + 100, img_size.height + 100))

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

        f_style = ImageCrop.from_obj(graphic)
        assert f_style.prop_img_size.width > img_size_mm.width
        assert f_style.prop_img_size.height > img_size_mm.height

        # test image scale only
        style = ImageCrop(img_scale=Size(95, 90))
        style.apply(graphic)

        f_style = ImageCrop.from_obj(graphic)
        assert f_style.prop_img_size.width < img_size_mm.width
        assert f_style.prop_img_size.height < img_size_mm.height

        # test scale and image size. Expect scale to be ignored.
        style = ImageCrop(
            img_scale=Size(95, 90), img_size=SizeMM.from_mm100(img_size.width + 100, img_size.height + 100)
        )
        style.apply(graphic)

        f_style = ImageCrop.from_obj(graphic)
        # if scale dictates that the numbers should be biggger but
        # scale is ignore by design when crop is not present and
        # image size is present.
        assert f_style.prop_img_size.width > img_size_mm.width
        assert f_style.prop_img_size.height > img_size_mm.height

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
