from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.format.direct.para.area.img import (
    Img,
    ImageKind,
    ImgStyleKind,
    SizeMM,
    SizePercent,
    Offset,
    OffsetColumn,
    OffsetRow,
    RectanglePoint,
)
from ooodev.utils.images_lo import ImagesLo, BitmapArgs

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service
    from com.sun.star.style import ParagraphProperties  # service


def test_write_presets(loader, para_text) -> None:
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        cursor = Write.get_cursor(doc)

        presets = (
            ImageKind.PAINTED_WHITE,
            ImageKind.PAPER_TEXTURE,
            ImageKind.PAPER_CRUMPLED,
            ImageKind.PARCHMENT_PAPER,
            ImageKind.FENCE,
            ImageKind.INVOICE_PAPER,
            ImageKind.MAPLE_LEAVES,
            ImageKind.LAWN,
            ImageKind.COLORFUL_PEBBLES,
            ImageKind.COFFEE_BEANS,
            ImageKind.LITTLE_CLOUDS,
            ImageKind.BATHROOM_TILES,
            ImageKind.CONCRETE,
            ImageKind.PAPER_GRAPH,
            ImageKind.ZEBRA,
            ImageKind.WALL_OF_ROCK,
            ImageKind.BRICK_WALL,
            ImageKind.STONE_WALL,
            ImageKind.FLORAL,
            ImageKind.SPACE,
            ImageKind.COLOR_STRIPES,
            ImageKind.ICE_LIGHT,
            ImageKind.POOL,
            ImageKind.MARBLE,
            ImageKind.SAND_LIGHT,
            ImageKind.STONE,
            ImageKind.WHITE_DIFFUSION,
            ImageKind.SURFACE,
            ImageKind.CARDBOARD,
            ImageKind.STUDIO,
        )
        cursor_p = Write.get_paragraph_cursor(cursor)
        for preset in presets:
            img = Img.from_preset(preset)

            Write.append_para(cursor=cursor, text=para_text, styles=(img,))

            cursor_p.gotoPreviousParagraph(False)
            cursor_p.gotoStartOfParagraph(False)
            cursor_p.gotoEndOfParagraph(True)
            fp = cast("FillProperties", cursor_p.TextParagraph)
            # point = preset._get_point()
            # assert fp.FillBitmapSizeX == point.x
            # assert fp.FillBitmapSizeY == point.y
            assert fp.FillBitmapName == str(preset)
            assert fp.FillBitmap is not None
            assert fp.FillBitmapMode == ImgStyleKind.TILED.value
            pp = cast("ParagraphProperties", cursor_p)
            assert pp.ParaBackColor == -1
            assert pp.ParaBackTransparent == True

            cursor_p.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_write_image(loader, para_text, fix_image_path) -> None:
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        fnm = fix_image_path("img_brick.png")
        bargs = BitmapArgs("Bitmap ", auto_name=True)
        xb = ImagesLo.get_bitmap(fnm, bargs)
        img = Img(
            bitmap=xb,
            name=bargs.out_name,
            mode=ImgStyleKind.TILED,
            size=SizeMM(67.73, 67.73),
            position=RectanglePoint.MIDDLE_MIDDLE,
            pos_offset=Offset(0, 0),
            tile_offset=OffsetRow(0),
            auto_name=False,
        )

        cursor = Write.get_cursor(doc)
        cursor_p = Write.get_paragraph_cursor(cursor)

        Write.append_para(cursor=cursor, text=para_text, styles=(img,))

        cursor_p.gotoPreviousParagraph(False)
        cursor_p.gotoStartOfParagraph(False)
        cursor_p.gotoEndOfParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillBitmapName == bargs.out_name
        assert fp.FillBitmap is not None
        assert fp.FillBitmapMode == ImgStyleKind.TILED.value
        pp = cast("ParagraphProperties", cursor_p)
        assert pp.ParaBackColor == -1
        assert pp.ParaBackTransparent == True

        cursor_p.gotoEnd(False)

        # Test again this time no bitmap just name
        img = Img(
            name=bargs.out_name,
            mode=ImgStyleKind.TILED,
            size=SizeMM(67.73, 67.73),
            position=RectanglePoint.MIDDLE_MIDDLE,
            pos_offset=Offset(0, 0),
            tile_offset=OffsetRow(0),
            auto_name=False,
        )

        Write.append_para(cursor=cursor, text=para_text, styles=(img,))

        cursor_p.gotoPreviousParagraph(False)
        cursor_p.gotoStartOfParagraph(False)
        cursor_p.gotoEndOfParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillBitmapName == bargs.out_name
        assert fp.FillBitmap is not None
        assert fp.FillBitmapMode == ImgStyleKind.TILED.value
        pp = cast("ParagraphProperties", cursor_p)
        assert pp.ParaBackColor == -1
        assert pp.ParaBackTransparent == True

        cursor_p.gotoEnd(False)

        Lo.delay(delay)

    finally:
        Lo.close_doc(doc)
