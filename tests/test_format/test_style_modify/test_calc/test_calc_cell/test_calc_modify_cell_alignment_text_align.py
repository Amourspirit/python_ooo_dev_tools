from __future__ import annotations
import pytest


if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format import Styler
from ooodev.format.calc.modify.cell.alignment import TextAlign, VertAlignKind, HoriAlignKind, StyleCellKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.calc import Calc
from ooodev.units.unit_mm100 import UnitMM100


def test_write(loader) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Calc.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        sheet = Calc.get_active_sheet()

        cell_obj = Calc.get_cell_obj("A1")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)

        style = TextAlign(
            hori_align=HoriAlignKind.CENTER, vert_align=VertAlignKind.MIDDLE, style_name=StyleCellKind.DEFAULT
        )
        Styler.apply(doc, style)

        f_style = TextAlign.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_hori_align == HoriAlignKind.CENTER
        assert f_style.prop_inner.prop_vert_align == VertAlignKind.MIDDLE
        assert f_style.prop_inner.prop_indent.get_value_mm100() == 0

        assert f_style.prop_inner.prop_hori_align == style.prop_inner.prop_hori_align
        assert f_style.prop_inner.prop_vert_align == style.prop_inner.prop_vert_align

        # ====================================================================

        cell_obj = Calc.get_cell_obj("B1")
        Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj)

        indent = UnitMM100.from_pt(10)
        style = TextAlign(
            hori_align=HoriAlignKind.LEFT,
            indent=indent,
            vert_align=VertAlignKind.TOP,
            style_name=StyleCellKind.DEFAULT,
        )
        Styler.apply(doc, style)

        f_style = TextAlign.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_hori_align == HoriAlignKind.LEFT
        assert f_style.prop_inner.prop_vert_align == VertAlignKind.TOP
        assert f_style.prop_inner.prop_indent.get_value_mm100() in range(indent.value - 2, indent.value + 3)  # +- 2

        assert f_style.prop_inner.prop_hori_align == style.prop_inner.prop_hori_align
        assert f_style.prop_inner.prop_vert_align == style.prop_inner.prop_vert_align

        # ====================================================================

        cell_obj = Calc.get_cell_obj("A2")
        Calc.set_val(value="distributed text", sheet=sheet, cell_obj=cell_obj)

        style = TextAlign(
            hori_align=HoriAlignKind.DISTRIBUTED,
            vert_align=VertAlignKind.DISTRIBUTED,
            style_name=StyleCellKind.DEFAULT,
        )
        Styler.apply(doc, style)

        f_style = TextAlign.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_hori_align == HoriAlignKind.DISTRIBUTED
        assert f_style.prop_inner.prop_vert_align == VertAlignKind.DISTRIBUTED
        assert f_style.prop_inner.prop_indent.get_value_mm100() in range(indent.value - 2, indent.value + 3)  # +- 2

        assert f_style.prop_inner.prop_hori_align == style.prop_inner.prop_hori_align
        assert f_style.prop_inner.prop_vert_align == style.prop_inner.prop_vert_align

        # ====================================================================

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
