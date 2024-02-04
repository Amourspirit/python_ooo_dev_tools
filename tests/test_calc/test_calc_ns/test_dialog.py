from __future__ import annotations
import pytest
from typing import Any, cast
import uno
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.gbl_named_event import GblNamedEvent

if __name__ == "__main__":
    pytest.main([__file__])


def test_msg_box(loader) -> None:
    # get_sheet is overload method.
    # testing each overload.
    from ooodev.calc import CalcDoc
    from ooo.dyn.awt.message_box_results import MessageBoxResultsEnum

    def on_dialog_creating(source: Any, event_args: CancelEventArgs) -> None:
        event_args.cancel = True
        event_args.event_data["result"] = MessageBoxResultsEnum.OK

    doc = CalcDoc.create_doc(loader)
    try:
        doc.subscribe_event(GblNamedEvent.MSG_BOX_CREATING, on_dialog_creating)
        result = doc.msgbox("Test Message")
        assert result == MessageBoxResultsEnum.OK

    finally:
        doc.close_doc()


def test_input_box(loader) -> None:
    # get_sheet is overload method.
    # testing each overload.
    from ooodev.calc import CalcDoc

    def on_dialog_creating(source: Any, event_args: CancelEventArgs) -> None:
        event_args.cancel = True
        event_args.event_data["result"] = "Test Input"

    doc = CalcDoc.create_doc(loader)
    try:
        doc.subscribe_event(GblNamedEvent.INPUT_BOX_CREATING, on_dialog_creating)
        result = doc.input_box(title="Test Title", msg="Test Message")
        assert result == "Test Input"

    finally:
        doc.close_doc()
