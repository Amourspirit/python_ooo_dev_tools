from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno

from ooodev.format.writer.direct.obj.options import Names
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


def test_write(loader, formula_text) -> None:
    # testing prev and next is not currently possible.
    # see Names class and read comments for details.

    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)

        style = Names(name="skinner", desc="Skinner Pointing", alt="Pointer")

        content = Write.add_formula(cursor=cursor, formula=formula_text, styles=(style,))

        f_style = Names.from_obj(content)
        assert f_style.prop_name == style.prop_name
        assert f_style.prop_desc == style.prop_desc
        assert f_style.prop_alt == style.prop_alt

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
