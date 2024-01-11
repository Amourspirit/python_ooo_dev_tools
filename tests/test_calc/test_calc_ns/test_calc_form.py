from __future__ import annotations
import pytest
from typing import cast
import uno
from com.sun.star.container import XNamed

if __name__ == "__main__":
    pytest.main([__file__])


def test_form_get_index_by_name(loader) -> None:
    # get_sheet is overload method.
    # testing each overload.
    from ooodev.utils.lo import Lo
    from ooodev.calc import Calc
    from ooodev.calc import CalcDoc

    doc = CalcDoc(Calc.create_doc(loader))
    try:
        sheet = doc.sheets[0]
        dp = sheet.draw_page
        forms = dp.forms
        forms.add_form("test")
        assert len(forms) == 1
        form = forms[0]
        btn1 = form.insert_control_button(
            x=10,
            y=10,
            width=100,
            height=20,
            label="Button 1",
        )
        btn2 = form.insert_control_button(
            x=10,
            y=10,
            width=100,
            height=20,
            label="Button 2",
        )
        index = form.get_index_by_name(btn2.name)
        assert index == 1
        index = form.get_index_by_name(btn1.name)
        assert index == 0

        assert len(form) == 2

        for itm in form:
            assert cast(XNamed, itm).getName() in (btn1.name, btn2.name)

    finally:
        doc.close_doc()
