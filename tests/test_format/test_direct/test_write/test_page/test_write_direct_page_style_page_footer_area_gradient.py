from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.page.footer import Footer
from ooodev.format.writer.direct.page.footer.area import Gradient, PresetGradientKind
from ooodev.format.writer.direct.char.font import Font
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.color import StandardColor
from ooodev.utils.info import Info
from ooodev.units import UnitMM100


def test_write(loader, para_text) -> None:
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)
        if not Lo.bridge_connector.headless:
            Write.append_para(cursor=cursor, text=para_text)

        height = UnitMM100.from_mm(10.0)
        spacing = UnitMM100.from_mm(3.0)
        m_left = UnitMM100.from_mm(1.5)
        m_right = UnitMM100.from_mm(2.0)

        footer = Footer(
            on=True,
            shared_first=True,
            shared=True,
            height=height,
            spacing=spacing,
            spacing_dyn=True,
            margin_left=m_left,
            margin_right=m_right,
        )
        style = Gradient.from_preset(preset=PresetGradientKind.NEON_LIGHT)
        font = Font(b=True, color=StandardColor.BLUE, size=16)
        Write.set_footer(text_doc=doc, text="Footer", styles=[footer, font, style])

        props = Info.get_style_props(doc=doc, family_style_name="PageStyles", prop_set_nm="Standard")
        f_style = Gradient.from_obj(props)
        assert f_style.prop_inner == style.prop_inner

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
