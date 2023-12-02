from __future__ import annotations
import pytest
from typing import Any

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.calc import Calc
from ooodev.adapter.sheet.spreadsheet_document_comp import SpreadsheetDocumentComp
from ooodev.events.args.event_args import EventArgs


def test_spreadsheet_document_comp(loader) -> None:
    name = ""

    def on_event(src: Any, event: EventArgs, control_src: Any, *args, **kwargs) -> None:
        nonlocal name
        assert event is not None
        name = event.event_name

    doc = Calc.create_doc(loader=loader)
    assert doc is not None, "Could not create new document"
    comp = SpreadsheetDocumentComp(doc)  # type: ignore
    assert comp.component
    setting = comp.spreadsheet_document_settings
    assert setting

    office_doc = comp.office_document
    assert office_doc

    comp.add_event_modified(on_event)
    delay = 0
    if not Lo.bridge_connector.headless:
        GUI.set_visible(visible=True, doc=doc)
    sheet = Calc.get_sheet(doc=doc, idx=0)
    Calc.set_val(sheet=sheet, col=0, row=0, value="test")
    assert name == "modified"

    Lo.delay(delay)
    Lo.close(doc)  # type: ignore
