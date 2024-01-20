from __future__ import annotations
from pathlib import Path
from typing import Any
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.calc import Calc, CalcDoc
from ooodev.calc import CalcNamedEvent
from ooodev.events.args.cancel_event_args_export import CancelEventArgsExport
from ooodev.events.args.event_args_export import EventArgsExport
from ooodev.calc.filter.export_jpg import ExportJpgT
from ooodev.calc.filter.export_png import ExportPngT


def test_export_range_jpg(loader, tmp_path) -> None:
    quality = 0

    def on_exporting(source: Any, args: CancelEventArgsExport[ExportJpgT]) -> None:
        # args.event_data["pixel_width"] = 1192
        # args.event_data["pixel_height"] = 1673
        # args.event_data["logical_width"] = 27522
        # args.event_data["logical_height"] = 38628
        args.event_data["quality"] = 80

    def on_exported(source: Any, args: EventArgsExport[ExportJpgT]) -> None:
        nonlocal quality
        quality = args.event_data["quality"]

    doc = CalcDoc(Calc.create_doc(loader))
    try:
        pth = Path(tmp_path, "test_export_range.jpg")

        vals = (
            ("", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"),
            ("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3, -67.3, 30.5, 23.2, -97.3, 22.4, 23.5),
            ("Jones", 21, 40.9, -57.5, -23.4, 34.5, 59.3, 27.3, -38.5, 43.2, 57.3, 25.4, 28.5),
            ("Brown", 31.45, -20.9, -117.5, 23.4, -114.5, 115.3, -171.3, 89.5, 41.2, 71.3, 25.4, 38.5),
        )
        sheet = doc.sheets[0]
        sheet.set_array(values=vals, name="A1")
        rng = sheet.get_range(range_name="A1:M4")
        # rng.export_as_image(pth)
        rng.subscribe_event(CalcNamedEvent.EXPORTING_RANGE_JPG, on_exporting)
        rng.subscribe_event(CalcNamedEvent.EXPORTED_RANGE_JPG, on_exported)
        rng.export_jpg(fnm=pth, resolution=96)
        assert quality == 80
        assert pth.exists()

        # filter_names = Info.get_filter_names()
        # assert filter_names
        # fp = Info.get_filter_props("calc_jpg_Export")
        # assert fp

    finally:
        doc.close_doc()


def test_export_range_png(loader, tmp_path) -> None:
    compression = 0

    def on_exporting(source: Any, args: CancelEventArgsExport[ExportPngT]) -> None:
        args.event_data["compression"] = 8

    def on_exported(source: Any, args: EventArgsExport[ExportPngT]) -> None:
        nonlocal compression
        compression = args.event_data["compression"]

    doc = CalcDoc(Calc.create_doc(loader))
    try:
        pth = Path(tmp_path, "test_export_range.png")

        vals = (
            ("", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"),
            ("Smith", 42, 58.9, -66.5, 43.4, 44.5, 45.3, -67.3, 30.5, 23.2, -97.3, 22.4, 23.5),
            ("Jones", 21, 40.9, -57.5, -23.4, 34.5, 59.3, 27.3, -38.5, 43.2, 57.3, 25.4, 28.5),
            ("Brown", 31.45, -20.9, -117.5, 23.4, -114.5, 115.3, -171.3, 89.5, 41.2, 71.3, 25.4, 38.5),
        )
        sheet = doc.sheets[0]
        sheet.set_array(values=vals, name="A1")
        rng = sheet.get_range(range_name="A1:M4")
        # rng.export_as_image(pth)
        rng.subscribe_event(CalcNamedEvent.EXPORTING_RANGE_PNG, on_exporting)
        rng.subscribe_event(CalcNamedEvent.EXPORTED_RANGE_PNG, on_exported)
        rng.export_png(fnm=pth, resolution=300)
        assert compression == 8
        assert pth.exists()

    finally:
        doc.close_doc()
