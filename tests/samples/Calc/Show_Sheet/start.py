#!/usr/bin/env python
# coding: utf-8
#
# on wayland (some versions of Linux)
# may get error:
#    (soffice:67106): Gdk-WARNING **: 02:35:12.168: XSetErrorHandler() called with a GDK error trap pushed. Don't do that.
# This seems to be a Wayland/Java compatability issues.
# see: http://www.babelsoft.net/forum/viewtopic.php?t=24545

import sys
import argparse
from typing import Any, cast

from ooodev.dialog.msgbox import MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
from ooodev.dialog.input import Input
from ooodev.events.args.cancel_event_args import CancelEventArgs
from ooodev.events.gbl_named_event import GblNamedEvent
from ooodev.events.lo_events import LoEvents
from ooodev.office.calc import Calc
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo


from com.sun.star.util import XProtectable


def args_add(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "-f",
        "--file",
        help="File path of input file to convert",
        action="store",
        dest="file_path",
        required=True,
    )
    parser.add_argument(
        "-r", "--read-only", help="Read only mode", action="store_true", dest="read_only", default=False
    )
    parser.add_argument("-s", "--show", help="Show Document", action="store_true", dest="show", default=False)
    parser.add_argument("-v", "--verbose", help="Verbose output", action="store_true", dest="verbose", default=False)


def on_lo_print(source: Any, e: CancelEventArgs) -> None:
    # this method is a callback for ooodev internal printing
    # by setting e.canecl = True all internal printing of ooodev is suppressed
    e.cancel = True


def main() -> int:
    # create parser to read terminal input
    parser = argparse.ArgumentParser(description="main")

    # add args to parser
    args_add(parser=parser)

    if len(sys.argv) <= 1:
        parser.print_help()
        return 0

    # read the current command line args
    args = parser.parse_args()

    visible = args.show
    if visible:
        delay = 2_000
    else:
        delay = 0

    if not args.verbose:
        # hook ooodev internal printing event
        LoEvents().on(GblNamedEvent.PRINTING, on_lo_print)

    loader = Lo.load_office(Lo.ConnectSocket())

    fnm = cast(str, args.file_path)

    try:
        doc = Calc.open_doc(fnm=fnm, loader=loader)

        if visible:
            GUI.set_visible(is_visible=visible, odoc=doc)

        Calc.goto_cell(cell_name="A1", doc=doc)
        sheet_names = Calc.get_sheet_names(doc=doc)
        print(f"Names of Sheets ({len(sheet_names)}):")
        for name in sheet_names:
            print(f"  {name}")

        sheet = Calc.get_sheet(doc=doc, sheet_name="Sheet1")
        Calc.set_active_sheet(doc=doc, sheet=sheet)
        pro = Lo.qi(XProtectable, sheet, True)
        pro.protect("foobar")
        print(f"Is protected: {pro.isProtected()}")

        Lo.delay(2000)
        pwd = Input.get_input("Password", "Enter sheet Password", is_password=True)
        if pwd == "foobar":
            pro.unprotect(pwd)
            MsgBox.msgbox("Password is Correct", "Password", boxtype=MessageBoxType.INFOBOX)
        else:
            MsgBox.msgbox("Password is incorrect", "Password", boxtype=MessageBoxType.ERRORBOX)

        msg_result = MsgBox.msgbox(
            "Do you wish to close document?",
            "All done",
            boxtype=MessageBoxType.QUERYBOX,
            buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
        )
        if msg_result == MessageBoxResultsEnum.YES:
            Lo.close_doc(doc=doc, deliver_ownership=True)
            Lo.close_office()
        else:
            print("Keeping document open")
    except Exception:
        Lo.close_office()
        raise

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
