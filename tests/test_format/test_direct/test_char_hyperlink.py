from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.char.hyperlink import Hyperlink, TargetKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


def test_hyperlink_props() -> None:
    ln_name = "ODEV"
    ln_url = "https://python-ooo-dev-tools.readthedocs.io/en/latest/"

    hl = Hyperlink(name=ln_name, url=ln_url)

    assert hl.prop_name == ln_name
    assert hl.prop_url == ln_url
    assert hl.prop_target == TargetKind.NONE
    assert hl.prop_visited_style == "Visited Internet Link"
    assert hl.prop_unvisited_style == "Internet link"

    hl.prop_name = "test"
    assert hl.prop_name == "test"

    hl.prop_url = "ftp:///notreal"
    assert hl.prop_url == "ftp:///notreal"

    hl.prop_target = TargetKind.PARENT
    assert hl.prop_target == TargetKind.PARENT

    hl.prop_unvisited_style = "unvisited"
    assert hl.prop_unvisited_style == "unvisited"

    hl.prop_visited_style = "visited"
    assert hl.prop_visited_style == "visited"

    hl.prop_url = None
    assert hl.prop_url is None
    assert hl._get("HyperLinkURL") is None
    assert hl._has("HyperLinkURL") == False

    hl.prop_name = None
    assert hl.prop_name is None
    assert hl._get("HyperLinkName") is None
    assert hl._has("HyperLinkName") == False

    hl = Hyperlink()
    assert hl.prop_url is None
    assert hl.prop_name is None
    assert hl._has("HyperLinkURL") == False
    assert hl._has("HyperLinkName") == False


def test_hyperlink(loader) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 2_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        # note: If the  first word of the document is styled then it also will set the paragraph style
        # this will result in the style automatically being applied to new pargraphs as they are written.
        ln_name = "ODEV"
        ln_url = "https://python-ooo-dev-tools.readthedocs.io/en/latest/"
        cursor = Write.get_cursor(doc)
        Write.append(cursor, "Libreoffice ")
        hl = Hyperlink(name=ln_name, url=ln_url)
        Write.append(cursor, "OOO Development Tools", (hl,))

        cursor.gotoEnd(False)
        cursor.goLeft(21, True)
        cp = cast("CharacterProperties", cursor)
        assert cp.HyperLinkName == ln_name
        assert cp.HyperLinkTarget == ""
        assert cp.HyperLinkURL == ln_url
        assert cp.UnvisitedCharStyleName == "Internet link"
        assert cp.VisitedCharStyleName == "Visited Internet Link"
        cursor.gotoEnd(False)

        Write.append_para(cursor, " Docs")

        Write.append(cursor, "Source on Github ")

        pos = Write.get_position(cursor)

        ln_name = "ODEV_GITHUB"
        ln_url = "https://github.com/Amourspirit/python_ooo_dev_tools"
        hl = Hyperlink(name=ln_name, url=ln_url, target=TargetKind.BLANK)
        Write.append(cursor, "OOO Development Tools on Github", (hl,))

        cursor.gotoEnd(False)
        cursor.goLeft(31, True)
        cp = cast("CharacterProperties", cursor)
        assert cp.HyperLinkName == ln_name
        assert cp.HyperLinkTarget == "_blank"
        assert cp.HyperLinkURL == ln_url
        assert cp.UnvisitedCharStyleName == "Internet link"
        assert cp.VisitedCharStyleName == "Visited Internet Link"
        cursor.gotoEnd(False)

        Write.style_left(cursor=cursor, pos=pos, styles=(Hyperlink.empty,))

        cursor.gotoEnd(False)
        cursor.goLeft(31, True)
        cp = cast("CharacterProperties", cursor)
        assert cp.HyperLinkName == ""
        assert cp.HyperLinkTarget == ""
        assert cp.HyperLinkURL == ""
        assert cp.UnvisitedCharStyleName == "Internet link"
        assert cp.VisitedCharStyleName == "Visited Internet Link"
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
