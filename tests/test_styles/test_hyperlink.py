from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.styles.char.hyperlink import Hyperlink, TargetKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo


if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


def test_hyperlink_props() -> None:
    ln_name = "ODEV"
    ln_url = "https://python-ooo-dev-tools.readthedocs.io/en/latest/"

    hl = Hyperlink(name=ln_name, url=ln_url)

    assert hl.name == ln_name
    assert hl.url == ln_url
    assert hl.target == TargetKind.NONE
    assert hl.visited_style == "Visited Internet Link"
    assert hl.unvisited_style == "Internet link"

    hl.name = "test"
    assert hl.name == "test"

    hl.url = "ftp:///notreal"
    assert hl.url == "ftp:///notreal"

    hl.target = TargetKind.PARENT
    assert hl.target == TargetKind.PARENT

    hl.unvisited_style = "unvisited"
    assert hl.unvisited_style == "unvisited"

    hl.visited_style = "visited"
    assert hl.visited_style == "visited"

    hl.url = None
    assert hl.url is None
    assert hl._get("HyperLinkURL") is None
    assert hl._has("HyperLinkURL") == False

    hl.name = None
    assert hl.name is None
    assert hl._get("HyperLinkName") is None
    assert hl._has("HyperLinkName") == False

    hl = Hyperlink()
    assert hl.url is None
    assert hl.name is None
    assert hl._has("HyperLinkURL") == False
    assert hl._has("HyperLinkName") == False


def test_hyperlink(loader, test_headless) -> None:
    delay = 0 if test_headless else 5_000
    from ooodev.office.write import Write

    doc = Write.create_doc()
    if not test_headless:
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
