#!/usr/bin/env python

import logging
from ooodev.loader.lo import Lo
from ooodev.calc import CalcDoc
from ooodev.loader.inst.options import Options
from ooodev.globals import GblEvents


def on_select_range(src, event):
    print("Selected Range From Event", event.event_data.rng_obj)


def on_before_select_range(src, event):
    print("Before Selected From Event", event.event_data.title)


def main():
    with Lo.Loader(connector=Lo.ConnectSocket(), opt=Options(log_level=logging.DEBUG)) as loader:
        doc = None
        try:

            doc = CalcDoc.create_doc(loader=loader, visible=True)
            GblEvents().subscribe_event("GlobalCalcRangeSelector", on_select_range)
            doc.subscribe_event("BeforePopupRangeSelection", on_before_select_range)
            doc.subscribe_event("AfterPopupRangeSelection", on_select_range)
            # doc.get_range_selection_from_popup()
            doc.insert_sheet("mysheet")
            sheet = doc.sheets[1]
            Lo.delay(1000)
            sheet.set_active()
            # sheet.invoke_range_selection()
            rng = sheet.get_range_selection_from_popup()
            print(rng)
            print("Done")

        finally:
            if doc is not None:
                doc.close()


if __name__ == "__main__":
    main()
