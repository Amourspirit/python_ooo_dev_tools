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


def test_shape_export_png(loader, tmp_path_fn) -> None:
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
        slide.height = slide.width  # slide.height * 2
        circle1 = slide.draw_circle(100, 100, 15)
        rect1 = slide.draw_rectangle(10, 10, 20, 20)
        rect1.apply_styles(shadow)

        resolution = 100
        img_path = Path(tmp_path_fn, f"draw_{rect1.name}_{resolution}.png")
        rect1.subscribe_event_shape_png_exporting(on_exporting)
        rect1.subscribe_event_shape_png_exported(on_exported)
        rect1.export_shape_png(fnm=img_path, resolution=resolution)
        assert img_path.exists()
        assert compression > 0
        assert url != ""

        compression = 0
        url = ""
        img_path = Path(tmp_path_fn, f"draw_{circle1.name}_{resolution}.png")
        circle1.subscribe_event_shape_png_exporting(on_exporting)
        circle1.subscribe_event_shape_png_exported(on_exported)
        circle1.export_shape_png(fnm=img_path, resolution=resolution)
        assert img_path.exists()
        assert compression > 0
        assert url != ""
    finally:
        doc.close_doc()


def test_shape_export_jpg(loader, tmp_path_fn) -> None:
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
        circle1 = slide.draw_circle(100, 100, 15)
        rect1 = slide.draw_rectangle(10, 10, 20, 20)
        rect1.apply_styles(shadow)

        resolution = 100
        img_path = Path(tmp_path_fn, f"draw_{rect1.name}_{resolution}.jpg")
        rect1.subscribe_event_shape_jpg_exporting(on_exporting)
        rect1.subscribe_event_shape_jpg_exported(on_exported)
        rect1.export_shape_jpg(fnm=img_path, resolution=resolution)
        assert img_path.exists()
        assert quality > 0
        assert url != ""

        quality = 0
        url = ""
        img_path = Path(tmp_path_fn, f"draw_{circle1.name}_{resolution}.jpg")
        circle1.subscribe_event_shape_jpg_exporting(on_exporting)
        circle1.subscribe_event_shape_jpg_exported(on_exported)
        circle1.export_shape_jpg(fnm=img_path, resolution=resolution)
        assert img_path.exists()
        assert quality > 0
        assert url != ""
    finally:
        doc.close_doc()


def test_shape_clone_export_jpg(loader, tmp_path_fn) -> None:
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
        circle1 = slide.draw_circle(100, 100, 15)
        rect1 = slide.draw_rectangle(10, 10, 20, 20)
        rect1.apply_styles(shadow)
        rect_copy = rect1.clone()

        resolution = 100
        img_path = Path(tmp_path_fn, f"draw_{rect_copy.name}_{resolution}.jpg")
        rect_copy.subscribe_event_shape_jpg_exporting(on_exporting)
        rect_copy.subscribe_event_shape_jpg_exported(on_exported)
        rect_copy.export_shape_jpg(fnm=img_path, resolution=resolution)
        assert img_path.exists()
        assert quality > 0
        assert url != ""

        quality = 0
        url = ""
        circle_copy = circle1.clone()
        img_path = Path(tmp_path_fn, f"draw_{circle_copy.name}_{resolution}.jpg")
        circle_copy.subscribe_event_shape_jpg_exporting(on_exporting)
        circle_copy.subscribe_event_shape_jpg_exported(on_exported)
        circle_copy.export_shape_jpg(fnm=img_path, resolution=resolution)
        assert img_path.exists()
        assert quality > 0
        assert url != ""
    finally:
        doc.close_doc()
