from __future__ import annotations
from typing import Any
import pytest
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])
import uno

from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.write import Write, WriteDoc
from ooodev.write import WriteNamedEvent
from ooodev.write.filter.export_jpg import ExportJpgT
from ooodev.write.filter.export_png import ExportPngT


def test_save_pages_png(loader, copy_fix_writer, tmp_path_fn):
    test_doc = copy_fix_writer("writer_dummy.odt")
    doc = WriteDoc(Write.open_doc(fnm=test_doc, loader=loader))

    compression = 0
    url = ""

    def on_exporting(source: Any, args: CancelEventArgsExport[ExportPngT]) -> None:
        nonlocal compression
        compression = args.event_data["compression"]

    def on_exported(source: Any, args: EventArgsExport[ExportPngT]) -> None:
        nonlocal url
        url = args.get("url")

    try:
        view = doc.get_view_cursor()

        doc_path = Path(doc.get_doc_path())  # type: ignore
        view.jump_to_last_page()
        view.subscribe_event(WriteNamedEvent.EXPORTED_PAGE_PNG, on_exported)
        view.subscribe_event(WriteNamedEvent.EXPORTING_PAGE_PNG, on_exporting)

        for i in range(view.get_page(), 0, -1):
            img_path = Path(tmp_path_fn, f"{doc_path.stem}_{i}.png")
            view.export_page_png(fnm=img_path)
            assert img_path.exists()
            assert compression > 0
            assert url != ""
            view.jump_to_previous_page()
            compression = 0
            url = ""

        view.jump_to_first_page()
    finally:
        doc.close_doc()


def test_save_pages_jpg(loader, copy_fix_writer, tmp_path_fn):
    test_doc = copy_fix_writer("writer_dummy.odt")
    doc = WriteDoc(Write.open_doc(fnm=test_doc, loader=loader))

    quality = 0
    url = ""

    def on_exporting(source: Any, args: CancelEventArgsExport[ExportJpgT]) -> None:
        nonlocal quality
        quality = args.event_data["quality"]

    def on_exported(source: Any, args: EventArgsExport[ExportJpgT]) -> None:
        nonlocal url
        url = args.get("url")

    try:
        view = doc.get_view_cursor()

        doc_path = Path(doc.get_doc_path())  # type: ignore
        view.jump_to_last_page()
        view.subscribe_event(WriteNamedEvent.EXPORTED_PAGE_JPG, on_exported)
        view.subscribe_event(WriteNamedEvent.EXPORTING_PAGE_JPG, on_exporting)

        for i in range(view.get_page(), 0, -1):
            img_path = Path(tmp_path_fn, f"{doc_path.stem}_{i}.jpg")
            view.export_page_jpg(fnm=img_path)
            assert img_path.exists()
            assert quality > 0
            assert url != ""
            view.jump_to_previous_page()
            quality = 0
            url = ""

        view.jump_to_first_page()
    finally:
        doc.close_doc()
