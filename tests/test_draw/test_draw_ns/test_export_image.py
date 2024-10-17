from __future__ import annotations
from typing import Any
import pytest
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.draw import Draw, DrawDoc
from ooodev.draw import DrawNamedEvent
from ooodev.draw.filter.export_png import ExportPngT

# from ooodev.draw.export.page_jpg import PageJpg
from ooodev.draw.export.page_png import PagePng


def test_save_pages_png(loader, tmp_path_fn):
    doc = DrawDoc(Draw.create_draw_doc(loader=loader))

    compression = 0
    url = ""

    def on_exporting(source: Any, args: CancelEventArgsExport[ExportPngT]) -> None:
        nonlocal compression
        compression = args.event_data["compression"]

    def on_exported(source: Any, args: EventArgsExport[ExportPngT]) -> None:
        nonlocal url
        url = args.get("url")

    try:
        slide = doc.slides[0]
        _ = slide.draw_circle(10, 10, 15)
        _ = slide.draw_rectangle(40, 20, 15, 15)

        exporter = PagePng(slide)
        exporter.subscribe_event(DrawNamedEvent.EXPORTED_PAGE_PNG, on_exported)
        exporter.subscribe_event(DrawNamedEvent.EXPORTING_PAGE_PNG, on_exporting)

        # doc_path = Path(doc.get_doc_path())  # type: ignore
        img_path = Path(tmp_path_fn, "image.jpg")

        exporter.export(fnm=img_path, resolution=96)
        assert img_path.exists()
        assert compression > 0
        assert url != ""

    finally:
        doc.close_doc()


# def test_save_pages_jpg(loader, tmp_path_fn):
#     doc = DrawDoc(Draw.create_draw_doc(loader=loader))

#     quality = 0
#     url = ""

#     def on_exporting(source: Any, args: CancelEventArgsExport[ExportJpgT]) -> None:
#         nonlocal quality
#         quality = args.event_data["quality"]

#     def on_exported(source: Any, args: EventArgsExport[ExportJpgT]) -> None:
#         nonlocal url
#         url = args.get("url")

#     try:
#         slide = doc.slides[0]
#         _ = slide.draw_circle(10, 10, 15)
#         _ = slide.draw_rectangle(40, 20, 15, 15)

#         exporter = PageJpg(slide, False)
#         exporter.subscribe_event(DrawNamedEvent.EXPORTED_PAGE_JPG, on_exported)
#         exporter.subscribe_event(DrawNamedEvent.EXPORTING_PAGE_JPG, on_exporting)

#         # doc_path = Path(doc.get_doc_path())  # type: ignore
#         img_path = Path(tmp_path_fn, "image.jpg")

#         exporter.export(fnm=img_path, resolution=600)
#         assert img_path.exists()
#         assert quality > 0
#         assert url != ""

#     finally:
#         doc.close_doc()
