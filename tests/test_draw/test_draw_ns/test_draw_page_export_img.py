from __future__ import annotations
from typing import Any
from pathlib import Path
import pytest


if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.draw import Draw, DrawDoc
from ooodev.draw import ImpressDoc
from ooodev.format.draw.direct.shadow import Shadow, ShadowLocationKind
from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.draw.filter.export_png import ExportPngT, ExportPng
from ooodev.draw.filter.export_jpg import ExportJpgT, ExportJpg
from ooodev.draw import DrawNamedEvent


def test_slide_export_png(loader, tmp_path_fn) -> None:
    compression = 0
    url = ""

    def on_exporting(source: Any, args: CancelEventArgsExport[ExportPngT]) -> None:
        nonlocal compression
        compression = args.event_data["compression"]
        # When translucent is True, the page margins are excluded.
        # When translucent is False, the page margins are included.
        args.event_data["translucent"] = True
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


def test_slide_save_jpg(loader, tmp_path_fn) -> None:
    from ooodev.utils.images_lo import ImagesLo

    resolution = 100

    doc = DrawDoc.create_doc(loader)
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

        px_width, px_height = ImagesLo.get_dpi_width_height(
            width=slide.component.Width, height=slide.component.Height, resolution=resolution
        )
        dt = ExportJpg(
            color_mode=True,
            pixel_width=px_width,
            pixel_height=px_height,
            quality=80,
            logical_width=px_width,
            logical_height=px_height,
        )

        mime = ImagesLo.change_to_mime("jpeg")
        img_path = Path(tmp_path_fn, f"draw_image_{resolution}.jpg")
        slide.save_page(fnm=img_path, mime_type=mime, filter_data=dt.to_filter_dict())
        assert img_path.exists()
    finally:
        doc.close_doc()


def test_slide_save_png(loader, tmp_path_fn) -> None:
    from ooodev.utils.images_lo import ImagesLo

    resolution = 100

    doc = DrawDoc.create_doc(loader)
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

        px_width, px_height = ImagesLo.get_dpi_width_height(
            width=slide.component.Width, height=slide.component.Height, resolution=resolution
        )
        # When translucent is True, the page margins are excluded.
        # When translucent is False, the page margins are included.
        dt = ExportPng(
            pixel_width=px_width,
            pixel_height=px_height,
            logical_width=px_width,
            logical_height=px_height,
            compression=8,
            translucent=False,
            interlaced=False,
        )

        mime = ImagesLo.change_to_mime("png")
        img_path = Path(tmp_path_fn, f"draw_image_{resolution}.png")
        slide.save_page(fnm=img_path, mime_type=mime, filter_data=dt.to_filter_dict())
        assert img_path.exists()
    finally:
        doc.close_doc()


def test_impress_export_page_img(loader, copy_fix_presentation, tmp_path_fn):
    test_doc = copy_fix_presentation("algs.odp")

    doc = ImpressDoc.open_doc(fnm=test_doc, loader=loader)
    try:
        slide = doc.slides[0]
        img_path = Path(tmp_path_fn, "impress_image.png")
        slide.export_page_png(fnm=img_path, resolution=100)
        assert img_path.exists()

        img_path = Path(tmp_path_fn, "impress_image.jpg")
        slide.export_page_jpg(fnm=img_path, resolution=100)
        assert img_path.exists()
    finally:
        doc.close_doc()


def test_impress_new_context_export_page_img(loader, copy_fix_presentation, tmp_path_fn):
    test_doc = copy_fix_presentation("algs.odp")
    from ooodev.loader.lo import Lo

    lo_inst = Lo.create_lo_instance(Lo.options)

    doc = ImpressDoc.open_doc(fnm=test_doc, lo_inst=lo_inst)
    try:
        slide = doc.slides[0]
        img_path = Path(tmp_path_fn, "impress_image.png")
        slide.export_page_png(fnm=img_path, resolution=100)
        assert img_path.exists()

        img_path = Path(tmp_path_fn, "impress_image.jpg")
        slide.export_page_jpg(fnm=img_path, resolution=100)
        assert img_path.exists()
    finally:
        doc.close_doc()


def test_new_context_slide_save_png(loader, tmp_path_fn) -> None:
    from ooodev.utils.images_lo import ImagesLo
    from ooodev.loader.lo import Lo

    lo_inst = Lo.create_lo_instance(Lo.options)

    resolution = 100

    # create a doc using a different instance
    doc = DrawDoc.create_doc(lo_inst=lo_inst)

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

        px_width, px_height = ImagesLo.get_dpi_width_height(
            width=slide.component.Width, height=slide.component.Height, resolution=resolution
        )
        # When translucent is True, the page margins are excluded.
        # When translucent is False, the page margins are included.
        dt = ExportPng(
            pixel_width=px_width,
            pixel_height=px_height,
            logical_width=px_width,
            logical_height=px_height,
            compression=8,
            translucent=False,
            interlaced=False,
        )

        mime = ImagesLo.change_to_mime("png")
        img_path = Path(tmp_path_fn, f"draw_image_{resolution}.png")
        slide.save_page(fnm=img_path, mime_type=mime, filter_data=dt.to_filter_dict())
        assert img_path.exists()
    finally:
        doc.close_doc()
