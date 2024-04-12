import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.units.unit_convert import UnitLength


def test_generic_rect_convert(loader) -> None:
    from ooodev.utils.data_type.generic_unit_rect import GenericUnitRect
    from ooodev.units.unit_mm import UnitMM

    x = 10
    y = 10
    width = 100
    height = 100

    rect = GenericUnitRect(UnitMM(x), UnitMM(y), UnitMM(width), UnitMM(height))
    assert rect.x == x
    assert rect.y == y
    assert rect.width == width
    assert rect.height == height
    rect_100 = rect.convert_to(unit_length=UnitLength.MM100)
    assert rect_100.x == x * 100
    assert rect_100.y == y * 100
    assert rect_100.width == width * 100
    assert rect_100.height == height * 100

    rect_10 = rect.convert_to(unit_length=UnitLength.MM10)
    assert rect_10.x == x * 10
    assert rect_10.y == y * 10
    assert rect_10.width == width * 10
    assert rect_10.height == height * 10

    uno_rect = rect_10.get_uno_rectangle()
    assert uno_rect.X == x * 10
    assert uno_rect.Y == y * 10
    assert uno_rect.Width == width * 10
    assert uno_rect.Height == height * 10


def test_generic_unit_point_convert(loader) -> None:
    from ooodev.utils.data_type.generic_unit_point import GenericUnitPoint
    from ooodev.units.unit_mm import UnitMM

    x = 10
    y = 10

    point = GenericUnitPoint(UnitMM(x), UnitMM(y))
    assert point.x == x
    assert point.y == y
    unit_100 = point.convert_to(unit_length=UnitLength.MM100)
    assert unit_100.x == x * 100
    assert unit_100.y == y * 100

    unit_10 = point.convert_to(unit_length=UnitLength.MM10)
    assert unit_10.x == x * 10
    assert unit_10.y == y * 10
    p = unit_10.get_uno_point()
    assert p.X == x * 10
    assert p.Y == y * 10


def test_generic_unit_size_convert(loader) -> None:
    from ooodev.utils.data_type.generic_unit_size import GenericUnitSize
    from ooodev.units.unit_mm import UnitMM

    width = 100
    height = 100

    sz = GenericUnitSize(UnitMM(width), UnitMM(height))
    assert sz.width == width
    assert sz.height == height

    unit_100 = sz.convert_to(unit_length=UnitLength.MM100)
    assert unit_100.width == width * 100
    assert unit_100.height == height * 100

    unit_10 = sz.convert_to(unit_length=UnitLength.MM10)
    assert unit_10.width == width * 10
    assert unit_10.height == height * 10
    p = unit_10.get_uno_size()
    assert p.Width == width * 10
    assert p.Height == height * 10


def test_generic_size_pos_convert(loader) -> None:
    from ooodev.utils.data_type.generic_unit_size_pos import GenericUnitSizePos
    from ooodev.units.unit_mm import UnitMM

    x = 10
    y = 10
    width = 100
    height = 100

    rect = GenericUnitSizePos(UnitMM(x), UnitMM(y), UnitMM(width), UnitMM(height))
    assert rect.x == x
    assert rect.y == y
    assert rect.width == width
    assert rect.height == height

    rect_100 = rect.convert_to(unit_length=UnitLength.MM100)
    assert rect_100.x == x * 100
    assert rect_100.y == y * 100
    assert rect_100.width == width * 100
    assert rect_100.height == height * 100

    rect_10 = rect.convert_to(unit_length=UnitLength.MM10)
    assert rect_10.x == x * 10
    assert rect_10.y == y * 10
    assert rect_10.width == width * 10
    assert rect_10.height == height * 10

    uno_rect = rect_10.get_uno_rectangle()
    assert uno_rect.X == x * 10
    assert uno_rect.Y == y * 10
    assert uno_rect.Width == width * 10
    assert uno_rect.Height == height * 10
