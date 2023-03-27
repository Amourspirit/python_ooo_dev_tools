from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.frame.hyperlink import LinkTo, ImageMapOptions, TargetKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.units.unit_mm import UnitMM


def test_write(loader, para_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
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

        ln_name = "ODEV"
        ln_url = "https://python-ooo-dev-tools.readthedocs.io/en/latest/"
        img_style = ImageMapOptions(server_map=True)

        link_style = LinkTo(name=ln_name, url=ln_url, target=TargetKind.SELF)

        frame = Write.add_text_frame(
            cursor=cursor,
            ypos=UnitMM(10.2),
            text=para_text,
            width=UnitMM(60),
            height=UnitMM(40),
            styles=(link_style, img_style),
        )

        f_link_style = LinkTo.from_obj(frame)
        assert f_link_style.prop_name == link_style.prop_name
        assert f_link_style.prop_target == link_style.prop_target
        assert f_link_style.prop_url == link_style.prop_url

        f_img = ImageMapOptions.from_obj(frame)
        assert f_img.prop_server_map

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
