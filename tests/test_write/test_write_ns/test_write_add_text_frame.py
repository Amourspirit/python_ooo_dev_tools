from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from typing import cast, TYPE_CHECKING
import uno

from ooodev.format.writer.direct.frame.borders import Padding
from ooodev.format.writer.direct.frame.borders import Side, Sides, BorderLineKind, LineSize
from ooodev.format.writer.direct.frame.area import Color
from ooodev.format.writer.direct.frame.type import (
    Anchor,
    AnchorKind,
    Size,
    RelativeKind,
    RelativeSize,
    AbsoluteSize,
    Position,
    HoriOrient,
    VertOrient,
    RelHoriOrient,
    RelVertOrient,
    Horizontal,
    Vertical,
)
from ooodev.format.writer.direct.frame.wrap import Settings, WrapTextMode
from ooodev.utils.color import StandardColor
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.write import WriteDoc
from ooodev.units.unit_mm import UnitMM

if TYPE_CHECKING:
    from com.sun.star.text import TextFrame  # service


def test_write(loader, para_text) -> None:
    # Test adding text frame with styles, as well as test getting the text frame from the document.

    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = WriteDoc.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = doc.get_cursor()
        if not Lo.bridge_connector.headless:
            cursor.append_para(text=para_text)

        style_anchor = Anchor(anchor=AnchorKind.AT_PARAGRAPH)
        amt = 2.0
        style_padding = Padding(all=amt)
        style_color = Color(color=StandardColor.DEFAULT_BLUE)
        side = Side(line=BorderLineKind.SOLID, color=StandardColor.BLUE_DARK3, width=LineSize.MEDIUM)
        style_border = Sides(all=side)
        style_position = Position(
            horizontal=Horizontal(position=HoriOrient.CENTER, rel=RelHoriOrient.PARAGRAPH_AREA),
            vertical=Vertical(position=VertOrient.TOP, rel=RelVertOrient.PARAGRAPH_TEXT_AREA_OR_BASE_LINE),
            mirror_even=False,
            keep_boundaries=False,
        )
        style_size = Size(
            width=RelativeSize(size=90, kind=RelativeKind.PARAGRAPH),
            height=AbsoluteSize(50.2),
            auto_height=True,
            auto_width=False,
        )
        style_mode = Settings(mode=WrapTextMode.DYNAMIC)

        assert doc.text_frames is not None
        assert len(doc.text_frames) == 0

        frame = doc.text_frames.add_text_frame(
            ypos=UnitMM(10.2),
            text=para_text,
            width=UnitMM(60),
            height=UnitMM(40),
            styles=(style_anchor, style_mode, style_padding, style_color, style_border, style_position, style_size),
        )
        assert len(doc.text_frames) == 1

        assert doc.text_frames.has_elements()
        assert doc.text_frames.has_by_name("Frame1")
        frame0 = cast("TextFrame", frame.component)
        frame1 = cast("TextFrame", doc.text_frames.get_by_name("Frame1").component)
        frame1.getName() == frame0.getName()

        for frm in doc.text_frames:
            assert frm.name == "Frame1"

        frame.name = "Frame one"
        assert frame.name == "Frame one"

        Lo.delay(delay)
    finally:
        doc.close_doc()
