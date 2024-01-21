from __future__ import annotations
from typing import Any
from pathlib import Path
import pytest


if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.draw import Draw, DrawDoc
from ooodev.format.draw.direct.shadow import Shadow, ShadowLocationKind
from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.draw.filter.export_png import ExportPngT
from ooodev.draw.filter.export_jpg import ExportJpgT
from ooodev.draw import DrawNamedEvent


def test_slide_export_png(loader, tmp_path_fn) -> None:
    compression = 0
    url = ""

    def on_exporting(source: Any, args: CancelEventArgsExport[ExportPngT]) -> None:
        nonlocal compression
        compression = args.event_data["compression"]
        # args.event_data["translucent"] = False
        # args.event_data["pixel_width"] = 4096
        # args.event_data["pixel_height"] = 4096

    def on_exported(source: Any, args: EventArgsExport[ExportPngT]) -> None:
        nonlocal url
        url = args.get("url")

    doc = DrawDoc(Draw.create_draw_doc(loader))
    try:
        shadow = Shadow(use_shadow=True, location=ShadowLocationKind.BOTTOM_RIGHT, distance=0.5, blur=0.5)

        slide = doc.slides[0]
        # slide.component.BorderLeft = 0
        # slide.component.BorderRight = 0
        # slide.component.BorderTop = 0
        # slide.component.BorderBottom = 0
        # slide.width = slide.width * 2
        slide.height = slide.width  # slide.height * 2
        _ = slide.draw_circle(100, 100, 15)
        _ = slide.draw_rectangle(40, 20, 15, 15)
        rect1 = slide.draw_rectangle(10, 10, 20, 20)
        rect1.apply_styles(shadow)
        width = slide.width - (slide.border_left + slide.border_right)
        rect2 = slide.draw_rectangle(5, 110, width - 10, 30)
        rect2.apply_styles(shadow)

        slide.subscribe_event(DrawNamedEvent.EXPORTED_PAGE_PNG, on_exported)
        slide.subscribe_event(DrawNamedEvent.EXPORTING_PAGE_PNG, on_exporting)

        resolution = 100
        img_path = Path(tmp_path_fn, f"draw_image_{resolution}.png")
        slide.export_page_png(fnm=img_path, resolution=resolution)
        assert img_path.exists()
        assert compression > 0
        assert url != ""
    finally:
        doc.close_doc()


def test_slide_export_jpg(loader, tmp_path_fn) -> None:
    quality = 0
    url = ""

    def on_exporting(source: Any, args: CancelEventArgsExport[ExportJpgT]) -> None:
        nonlocal quality
        quality = args.event_data["quality"]
        # args.event_data["pixel_width"] = 4096
        # args.event_data["pixel_height"] = 4096

    def on_exported(source: Any, args: EventArgsExport[ExportJpgT]) -> None:
        nonlocal url
        url = args.get("url")

    doc = DrawDoc(Draw.create_draw_doc(loader))
    try:
        shadow = Shadow(use_shadow=True, location=ShadowLocationKind.BOTTOM_RIGHT, distance=0.5, blur=0.5)

        slide = doc.slides[0]
        slide.height = slide.width  # slide.height * 2
        _ = slide.draw_circle(100, 100, 15)
        _ = slide.draw_rectangle(40, 20, 15, 15)
        rect1 = slide.draw_rectangle(10, 10, 20, 20)
        rect1.apply_styles(shadow)
        width = slide.width - (slide.border_left + slide.border_right)
        rect2 = slide.draw_rectangle(5, 110, width - 10, 30)
        rect2.apply_styles(shadow)

        slide.subscribe_event(DrawNamedEvent.EXPORTED_PAGE_JPG, on_exported)
        slide.subscribe_event(DrawNamedEvent.EXPORTING_PAGE_JPG, on_exporting)

        resolution = 100
        img_path = Path(tmp_path_fn, f"draw_image_{resolution}.jpg")
        slide.export_page_jpg(fnm=img_path, resolution=resolution)
        assert img_path.exists()
        assert quality > 0
        assert url != ""
    finally:
        doc.close_doc()
