from __future__ import annotations
import pytest
from typing import cast, TYPE_CHECKING


if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format import Styler
from ooodev.format.calc.modify.cell.alignment import Properties, TextDirectionKind, StyleCellKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.calc import Calc

if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


def test_write(loader, para_text) -> None:
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
        cell = Calc.get_cell(sheet, cell_obj)

        style = Properties(
            wrap_auto=True, hyphen_active=True, direction=TextDirectionKind.PAGE, style_name=StyleCellKind.ACCENT
        )
        Styler.apply(doc, style)

        f_style = Properties.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_wrap_auto == style.prop_inner.prop_wrap_auto
        assert f_style.prop_inner.prop_hyphen_active == style.prop_inner.prop_hyphen_active
        assert f_style.prop_inner.prop_direction == style.prop_inner.prop_direction

        # ====================================================================

        style = Properties(
            wrap_auto=False, shrink_to_fit=True, direction=TextDirectionKind.LR_TB, style_name=StyleCellKind.ACCENT
        )
        Styler.apply(doc, style)

        f_style = Properties.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_wrap_auto == style.prop_inner.prop_wrap_auto
        assert f_style.prop_inner.prop_shirnk_to_fit == style.prop_inner.prop_shirnk_to_fit
        assert f_style.prop_inner.prop_direction == style.prop_inner.prop_direction
        assert f_style.prop_inner.prop_direction == TextDirectionKind.LR_TB

        # ====================================================================

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
