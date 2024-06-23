#!/usr/bin/env python

import logging
from ooodev.loader.lo import Lo
from ooodev.calc import CalcDoc
from ooodev.loader.inst.options import Options


def main():
    with Lo.Loader(connector=Lo.ConnectSocket(), opt=Options(log_level=logging.DEBUG)) as loader:
        doc = None
        try:

            doc = CalcDoc.create_doc(loader=loader, visible=True)
            rng = doc.get_range_selection_from_popup()
            print("Range Sel", rng)
            doc.insert_sheet("mysheet")
            sheet = doc.sheets[1]
            sheet.set_active()
            rng = sheet.get_range_selection_from_popup()
            print("Range Sel", rng)
            print("Done")

        finally:
            if doc is not None:
                doc.close()


if __name__ == "__main__":
    main()
