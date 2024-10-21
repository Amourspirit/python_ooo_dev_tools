from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.calc.direct.cell.numbers import Numbers, NumberFormatEnum, NumberFormatIndexEnum
from ooodev.format import Styler
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.units import UnitCM


def test_calc_number(loader) -> None:
    delay = 0
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    try:
        sheet = Calc.get_sheet(doc)
        if not Lo.bridge_connector.headless:
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_150_PERCENT)

        Calc.set_col_width(sheet=sheet, width=UnitCM(5), idx=0)
        cell_obj = Calc.get_cell_obj("A1")
        Calc.set_val(value=10.0, sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        num_style = Numbers(NumberFormatEnum.CURRENCY)
        Styler.apply(cell, num_style)
        f_num = Numbers.from_obj(cell)
        assert f_num.prop_format_key == num_style.prop_format_key

        cell_obj = Calc.get_cell_obj("A2")
        Calc.set_val(value=-123.0, sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        num_style = Numbers(num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2_RED)
        Styler.apply(cell, num_style)
        f_num = Numbers.from_obj(cell)
        assert f_num.prop_format_key == num_style.prop_format_key

        cell_obj = Calc.get_cell_obj("A3")
        Calc.set_val(value=0.000000034, sheet=sheet, cell_obj=cell_obj, styles=[num_style.scientific])
        cell = Calc.get_cell(sheet, cell_obj)
        f_num = Numbers.from_obj(cell)
        assert f_num.prop_format_key == num_style.scientific.prop_format_key

        num_style = Numbers(num_format_index=NumberFormatIndexEnum.SCIENTIFIC_000E000)
        cell_obj = Calc.get_cell_obj("A4")
        Calc.set_val(value=0.000000034, sheet=sheet, cell_obj=cell_obj, styles=[num_style])
        cell = Calc.get_cell(sheet, cell_obj)
        f_num = Numbers.from_obj(cell)
        assert f_num.prop_format_key == num_style.prop_format_key

        num_style = Numbers().date
        cell_obj = Calc.get_cell_obj("A5")
        Calc.set_val(value="03/17/2024", sheet=sheet, cell_obj=cell_obj, styles=[num_style])
        cell = Calc.get_cell(sheet, cell_obj)
        f_num = Numbers.from_obj(cell)
        assert f_num.prop_format_key == num_style.prop_format_key

        num_style = Numbers(num_format_index=NumberFormatIndexEnum.DATE_QQJJ)
        cell_obj = Calc.get_cell_obj("A6")
        Calc.set_val(value="03/17/2024", sheet=sheet, cell_obj=cell_obj, styles=[num_style])
        cell = Calc.get_cell(sheet, cell_obj)
        f_num = Numbers.from_obj(cell)
        assert f_num.prop_format_key == num_style.prop_format_key

        num_style = Numbers.from_str("#,##0.00_);(#,##0.00)")
        cell_obj = Calc.get_cell_obj("A7")
        Calc.set_val(value=1158.999, sheet=sheet, cell_obj=cell_obj, styles=[num_style])
        cell = Calc.get_cell(sheet, cell_obj)
        f_num = Numbers.from_obj(cell)
        assert f_num.prop_format_key == num_style.prop_format_key

        # user defined
        c_format = "#,##0.000"
        num_style = Numbers.from_str(c_format, auto_add=True)
        cell_obj = Calc.get_cell_obj("A8")
        Calc.set_val(value=1158.999, sheet=sheet, cell_obj=cell_obj, styles=[num_style])
        cell = Calc.get_cell(sheet, cell_obj)
        f_num = Numbers.from_obj(cell)
        key = num_style.prop_format_key
        assert num_style.prop_format_str == c_format
        assert f_num.prop_format_key == num_style.prop_format_key

        num_style = Numbers.from_index(key)
        cell_obj = Calc.get_cell_obj("A9")
        Calc.set_val(value=1158.999, sheet=sheet, cell_obj=cell_obj, styles=[num_style])
        cell = Calc.get_cell(sheet, cell_obj)
        f_num = Numbers.from_obj(cell)
        key = num_style.prop_format_key
        assert f_num.prop_format_key == num_style.prop_format_key
        assert f_num.prop_format_str == c_format

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
