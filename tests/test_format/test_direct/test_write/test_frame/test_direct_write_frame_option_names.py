from __future__ import annotations
import pytest
from typing import cast, TYPE_CHECKING

if __name__ == "__main__":
    pytest.main([__file__])

import uno

from ooodev.format.writer.direct.frame.options import Names
from ooodev.format.writer.direct.frame.type import (
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
    Anchor,
    AnchorKind,
)
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.office.write import Write
from ooodev.units.unit_mm import UnitMM


if TYPE_CHECKING:
    from com.sun.star.text import ChainedTextFrame

# see comments in ooodev.office.write.Write.add_text_frame() method.
# see comments in oodev.format.direct.frame.options.names.Names class.


def test_write_direct_chain(loader, para_text) -> None:
    # testing prev and next is not currently possible.
    # see Names class and read comments for details.

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

        style_anchor = Anchor(anchor=AnchorKind.AT_PAGE)
        size_style1 = Size(
            width=RelativeSize(size=30, kind=RelativeKind.PARAGRAPH),
            height=AbsoluteSize(118),
            auto_height=False,
            auto_width=False,
        )
        style_pos1 = Position(
            horizontal=Horizontal(position=HoriOrient.LEFT_OR_INSIDE, rel=RelHoriOrient.PAGE_TEXT_AREA),
            vertical=Vertical(position=VertOrient.TOP, rel=RelVertOrient.PAGE_TEXT_AREA),
            mirror_even=False,
        )

        style_n1 = Names(name="Frame_one", desc="Frame ONE")

        _ = Write.add_text_frame(
            cursor=cursor,
            text=para_text,
            page_num=0,
            styles=(
                style_anchor,
                size_style1,
                style_pos1,
                style_n1,
            ),
        )

        size_style2 = Size(
            width=RelativeSize(size=30, kind=RelativeKind.PARAGRAPH),
            height=AbsoluteSize(118),
            auto_height=False,
            auto_width=False,
        )
        style_pos2 = Position(
            horizontal=Horizontal(
                position=HoriOrient.FROM_LEFT_OR_INSIDE, rel=RelHoriOrient.PAGE_TEXT_AREA, amount=120.0
            ),
            vertical=Vertical(position=VertOrient.TOP, rel=RelVertOrient.PAGE_TEXT_AREA),
            mirror_even=False,
        )
        style_n2 = Names(name="Frame_two", prev=style_n1.prop_name)

        _ = Write.add_text_frame(
            cursor=cursor,
            page_num=0,
            styles=(
                style_anchor,
                size_style2,
                style_pos2,
                style_n2,
            ),
        )

        frames = Write.get_text_frames(doc)
        assert frames.hasByName(style_n1.prop_name)
        assert frames.hasByName(style_n2.prop_name)
        f1 = cast("ChainedTextFrame", frames.getByName(style_n1.prop_name))
        f2 = cast("ChainedTextFrame", frames.getByName(style_n2.prop_name))

        assert f1.ChainNextName == style_n2.prop_name
        assert f2.ChainPrevName == style_n1.prop_name

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_write(loader, para_text) -> None:
    # testing prev and next is not currently possible in headless mode.
    # see Names class and read comments for details.

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
            cursor=cursor, ypos=UnitMM(22.0), text=para_text, width=UnitMM(60), height=UnitMM(80), styles=(style_n1,)
        )

        frames = Write.get_text_frames(doc)
        assert frames.hasByName(style_n1.prop_name)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_write_frame_doc(loader, copy_fix_writer, para_text) -> None:
    delay = 0
    test_doc = copy_fix_writer("two_frame.odt")
    doc = Write.open_doc(fnm=test_doc)
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        style_n1 = Names(name="Frame_one")
        style_n2 = Names(name="Frame_two", prev=style_n1.prop_name)

        frames = Write.get_text_frames(doc)

        f1 = cast("ChainedTextFrame", frames.getByName(style_n1.prop_name))
        f2 = cast("ChainedTextFrame", frames.getByName(style_n2.prop_name))

        style_n2.apply(f2)
        assert f1.ChainNextName == style_n2.prop_name
        assert f2.ChainPrevName == style_n1.prop_name

        style_n1 = Names(name="Frame_first", desc="Frame First", next=style_n2.prop_name)
        style_n1.apply(f1)
        assert f1.ChainNextName == style_n2.prop_name
        assert f2.ChainPrevName == style_n1.prop_name
        assert f1.Description == "Frame First"

        style_n1.prop_next = ""
        style_n1.prop_prev = ""
        style_n1.apply(f1)
        assert f1.ChainNextName == ""
        assert f2.ChainPrevName == ""

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
