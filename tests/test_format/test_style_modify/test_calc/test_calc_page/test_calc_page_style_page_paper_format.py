from __future__ import annotations
from typing import cast, TYPE_CHECKING
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.calc.modify.page.page import PaperFormat, SizeMM, PaperFormatKind, CalcStylePageKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.calc import Calc

if TYPE_CHECKING:
    from com.sun.star.style import PageStyle


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

        presets = (
            PaperFormatKind.A3,
            PaperFormatKind.A4,
            PaperFormatKind.A5,
            PaperFormatKind.A6,
            PaperFormatKind.B4_ISO,
            PaperFormatKind.B4_JIS,
            PaperFormatKind.B5_ISO,
            PaperFormatKind.B5_JIS,
            PaperFormatKind.B6_ISO,
            PaperFormatKind.B6_JIS,
            PaperFormatKind.ENVELOPE_9,
            PaperFormatKind.ENVELOPE_10,
            PaperFormatKind.ENVELOPE_11,
            PaperFormatKind.ENVELOPE_12,
            PaperFormatKind.ENVELOPE_7_3X4,
            PaperFormatKind.LONG_BOND,
            PaperFormatKind.LEGAL,
            PaperFormatKind.KAI_32,
            PaperFormatKind.LETTER,
        )
        for preset in presets:
            style = PaperFormat.from_preset(preset=preset, landscape=False, style_name=CalcStylePageKind.DEFAULT)
            style.apply(doc)
            props = style.get_style_props(doc)
            ps = cast("PageStyle", props)
            assert ps.IsLandscape == False

            f_style = PaperFormat.from_style(
                doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
            )
            f_sz = f_style.prop_inner.prop_size.get_size_mm100()
            sz = style.prop_inner.prop_size.get_size_mm100()
            assert f_sz.height in range(sz.height - 2, sz.height + 3)
            assert f_sz.width in range(sz.width - 2, sz.width + 3)

        for preset in presets:
            style = PaperFormat.from_preset(preset=preset, landscape=True, style_name=CalcStylePageKind.DEFAULT)
            style.apply(doc)
            props = style.get_style_props(doc)
            ps = cast("PageStyle", props)
            assert ps.IsLandscape

            f_style = PaperFormat.from_style(
                doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
            )
            f_sz = f_style.prop_inner.prop_size.get_size_mm100()
            sz = style.prop_inner.prop_size.get_size_mm100()
            assert f_sz.height in range(sz.height - 2, sz.height + 3)
            assert f_sz.width in range(sz.width - 2, sz.width + 3)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
