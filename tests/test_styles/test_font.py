from ooodev.styles.font import (
    Font,
    StrikeOutKind,
    UnderlineKind,
    WeightKind,
    CharSetKnid,
    FamilyKind,
    SlantKind,
)
from ooodev.styles import CommonColor
from ooodev.utils.info import Info
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo


def test_font(loader) -> None:
    ft = Font(
        name=Info.get_font_general_name(),
        charset=CharSetKnid.SYSTEM,
        family=FamilyKind.MODERN,
        b=True,
        i=True,
        u=True,
        color=CommonColor.BLUE,
        strike=StrikeOutKind.BOLD,
        underine_color=CommonColor.AQUA,
        super_script=True,
    )
    assert ft.name == Info.get_font_general_name()
    assert ft.charset == CharSetKnid.SYSTEM
    assert ft.family == FamilyKind.MODERN
    assert ft.b
    assert ft.i
    assert ft.u
    assert ft.weight == WeightKind.BOLD
    assert ft.slant == SlantKind.ITALIC
    assert ft.underline == UnderlineKind.SINGLE
    assert ft.color == CommonColor.BLUE
    assert ft.strike == StrikeOutKind.BOLD
    assert ft.underine_color == CommonColor.AQUA
    assert ft.super_script

    ft = Font(weight=WeightKind.BOLD, underine=UnderlineKind.BOLDDASH, slant=SlantKind.OBLIQUE, sub_script=True)
    assert ft.weight == WeightKind.BOLD
    assert ft.underline == UnderlineKind.BOLDDASH
    assert ft.slant == SlantKind.OBLIQUE
    assert ft.sub_script


def test_font_cursor(loader, test_headless) -> None:
    delay = 5_000
    from ooodev.office.write import Write

    if test_headless:
        GUI.set_visible()
    doc = Write.create_doc()
    try:
        ft = Font(b=True)
        cursor = Write.get_cursor(doc)
        Write.append(cursor, "hello")
        cursor.goLeft(5, True)
        ft.apply_style(cursor)
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
