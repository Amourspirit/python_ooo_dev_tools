from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast
import random

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.draw import Draw

# from ooodev.format.inner.direct.write.fill.area.hatch import Hatch, PresetHatchKind
from ooodev.format.inner.direct.structs.hatch_struct import HatchStruct
from ooodev.format.draw.direct.area import Hatch, PresetHatchKind
from ooodev.utils.color import StandardColor
from ooodev.utils.table_helper import TableHelper

from ooo.dyn.drawing.fill_style import FillStyle
from ooo.dyn.drawing.hatch_style import HatchStyle

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service


def test_draw(loader) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Draw.create_draw_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        slide = Draw.get_slide(doc)

        width = 36
        height = 36
        x = width / 2
        y = height / 2

        rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
