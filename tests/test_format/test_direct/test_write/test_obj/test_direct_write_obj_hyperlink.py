from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.obj.hyperlink import LinkTo, ImageMapOptions, TargetKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write


def test_write(loader, formula_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)

        ln_name = "ODEV"
        ln_url = "https://python-ooo-dev-tools.readthedocs.io/en/latest/"
        img_style = ImageMapOptions(server_map=True)

        link_style = LinkTo(name=ln_name, url=ln_url, target=TargetKind.SELF)

        content = Write.add_formula(cursor=cursor, formula=formula_text, styles=(img_style, link_style))

        f_link_style = LinkTo.from_obj(content)
        assert f_link_style.prop_name == link_style.prop_name
        assert f_link_style.prop_target == link_style.prop_target
        assert f_link_style.prop_url == link_style.prop_url

        f_img = ImageMapOptions.from_obj(content)
        assert f_img.prop_server_map

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
