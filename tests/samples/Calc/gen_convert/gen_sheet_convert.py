from __future__ import annotations

from ooodev.units.convert.unit_length_kind import UnitLengthKind
from ooodev.units.convert.unit_area_kind import UnitAreaKind
from ooodev.loader import Lo
from ooodev.calc import CalcDoc

# Found a bug while working in this:
# https://bugs.documentfoundation.org/show_bug.cgi?id=160247
# The bug is that the CONVERT function does not work in Calc when set with setFormula.


def main():
    loader = Lo.load_office(Lo.ConnectSocket())
    doc = CalcDoc.create_doc(loader=loader, visible=True)
    try:
        vals = set()
        for val in UnitLengthKind:
            vals.add(val.value)
        if "m" in vals:
            vals.remove("m")
        data_col1 = []
        i = 0
        for val in vals:
            i += 1
            # data_col1.append([1, val, "m", f"=CONVERT(A{i}, B{i}, C{i})"])
            data_col1.append([1, val, "m"])
        sheet = doc.sheets[0]
        sheet.set_array(values=data_col1, name="A1")

        doc.sheets.insert_new_by_name("Area", -1)
        sheet = doc.sheets[-1]
        vals = set()
        for val in UnitAreaKind:
            vals.add(val.value)
        if "ar" in vals:
            vals.remove("ar")
        data_col1 = []
        i = 0
        for val in vals:
            i += 1
            data_col1.append([1, val, "ar"])
        sheet.set_array(values=data_col1, name="A1")

        # sheet["A1"].value = 2
        # sheet["B1"].value = 3  # "yd"
        # sheet["C1"].value = "m"
        # sheet["D1"].component.setFormula("=SUM(A1:B1)")
        # sheet["A2"].component.setFormula('=CONVERT(2, "yd", "m")')

        Lo.delay(500)
    finally:
        # doc.close()
        # Lo.close_office()
        print("done")


if __name__ == "__main__":
    main()
