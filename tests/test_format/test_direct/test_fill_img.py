from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.draw import Draw
from ooodev.format.draw.direct.area import (
    Img,
    PresetImageKind,
    ImgStyleKind,
)

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service


def test_draw(loader) -> None:
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Draw.create_draw_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        slide = Draw.get_slide(doc)
        width = 30
        height = 30
        max_col = 6
        presets = (
            PresetImageKind.PAINTED_WHITE,
            PresetImageKind.PAPER_TEXTURE,
            PresetImageKind.PAPER_CRUMPLED,
            PresetImageKind.PARCHMENT_PAPER,
            PresetImageKind.FENCE,
            PresetImageKind.INVOICE_PAPER,
            PresetImageKind.MAPLE_LEAVES,
            PresetImageKind.LAWN,
            PresetImageKind.COLORFUL_PEBBLES,
            PresetImageKind.COFFEE_BEANS,
            PresetImageKind.LITTLE_CLOUDS,
            PresetImageKind.BATHROOM_TILES,
            PresetImageKind.CONCRETE,
            PresetImageKind.PAPER_GRAPH,
            PresetImageKind.ZEBRA,
            PresetImageKind.WALL_OF_ROCK,
            PresetImageKind.BRICK_WALL,
            PresetImageKind.STONE_WALL,
            PresetImageKind.FLORAL,
            PresetImageKind.SPACE,
            PresetImageKind.COLOR_STRIPES,
            PresetImageKind.ICE_LIGHT,
            PresetImageKind.POOL,
            PresetImageKind.MARBLE,
            PresetImageKind.SAND_LIGHT,
            PresetImageKind.STONE,
            PresetImageKind.WHITE_DIFFUSION,
            PresetImageKind.SURFACE,
            PresetImageKind.CARDBOARD,
            PresetImageKind.STUDIO,
        )
        col = 0
        row = 1
        for preset in presets:
            col = col + 1
            if col > max_col:
                col = 1
                row += 1
            if col == 1:
                x = width / 2
            else:
                x = (width / 2) + (width * (col - 1))

            if row == 1:
                y = height / 2
            else:
                y = (height / 2) + (height * (row - 1))

            rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
            img = Img.from_preset(preset)
            img.apply(rec)
            fp = cast("FillProperties", rec)
            point = preset._get_point()

            assert fp.FillBitmapName == str(preset)
            assert fp.FillBitmap is not None
            assert fp.FillBitmapMode == ImgStyleKind.TILED.value
            assert fp.FillBitmapSizeX == point.x
            assert fp.FillBitmapSizeY == point.y

            # FillBitmapTile and FillBitmapStretch are not used anymore.
            # assert fp.FillBitmapTile == True
            # assert fp.FillBitmapStretch == False
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
