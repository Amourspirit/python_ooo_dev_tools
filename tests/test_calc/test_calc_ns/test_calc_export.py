from __future__ import annotations
from pathlib import Path
from typing import Any
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.calc import Calc, CalcDoc
from ooodev.utils.info import Info


def test_export_range_jpg(loader, tmp_path) -> None:
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
        rng.export_as_image(pth)
        assert pth.exists()

        # filter_names = Info.get_filter_names()
        # assert filter_names
        # fp = Info.get_filter_props("calc_jpg_Export")
        # assert fp

    finally:
        doc.close_doc()


def test_export_range_png(loader, tmp_path) -> None:
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
        rng.export_as_image(pth)
        assert pth.exists()

    finally:
        doc.close_doc()


def test_export_range_png_event(loader, tmp_path) -> None:
    from ooodev.events.args.cancel_event_args_generic import CancelEventArgsGeneric
    from ooodev.events.args.event_args_generic import EventArgsGeneric
    from ooodev.events.event_data.img_export_t import ImgExportT
    from ooodev.calc import CalcNamedEvent

    compression = 0

    def on_exporting(source: Any, args: CancelEventArgsGeneric[ImgExportT]) -> None:
        args.event_data["compression"] = 9

    def on_exported(source: Any, args: EventArgsGeneric[ImgExportT]) -> None:
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
        rng.subscribe_event(CalcNamedEvent.RANGE_EXPORTING_IMAGE, on_exporting)
        rng.subscribe_event(CalcNamedEvent.RANGE_EXPORTED_IMAGE, on_exported)
        rng.export_as_image(pth)
        assert pth.exists()
        assert compression == 9

    finally:
        doc.close_doc()
