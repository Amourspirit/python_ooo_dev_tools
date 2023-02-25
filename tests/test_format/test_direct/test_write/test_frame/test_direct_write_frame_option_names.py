from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno

# from com.sun.star.text import XTextFrame
# from com.sun.star.drawing import XShape

from ooodev.format.writer.direct.frame.options import Names
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.data_type.unit_mm import UnitMM


def test_write(loader, para_text) -> None:
    # testing prev and next is not currently possible.
    # see Names class and read comments for details.

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

        style_n1 = Names(name="Frame_one", desc="Frame ONE")

        _ = Write.add_text_frame(
            cursor=cursor, ypos=UnitMM(10.2), text=para_text, width=UnitMM(60), height=UnitMM(40), styles=(style_n1,)
        )

        frames = Write.get_text_frames(doc)
        assert frames.hasByName(style_n1.prop_name)

        # Raises error: see, https://bugs.documentfoundation.org/show_bug.cgi?id=153825
        # xframe = Lo.create_instance_msf(XTextFrame, "com.sun.star.text.ChainedTextFrame")
        # tf_shape = Lo.qi(XShape, xframe, True)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
