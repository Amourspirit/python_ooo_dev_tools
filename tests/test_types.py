import uno
from com.sun.star.frame import XComponentLoader

from ooodev.utils.info import Info


def test_detect_struct() -> None:
    from com.sun.star.awt import Point

    p1 = Point(10, 12)
    p2 = Point(10, 12)
    assert p1 is not None
    t = uno.getTypeByName(p1.typeName)
    assert t.typeClass.value == "STRUCT"
    assert p1.typeName == p2.typeName
    struct_attrs = [s for s in dir(p1.value) if not s.startswith("_")]
    for atr in struct_attrs:
        assert getattr(p1, atr) == getattr(p2, atr)


def test_ooo_detect_struct() -> None:
    from ooo.dyn.awt.point import Point

    p1 = Point(10, 12)
    p2 = Point(10, 12)
    assert p1 is not None
    t = uno.getTypeByName(p1.typeName)
    assert t.typeClass.value == "STRUCT"
    assert p1.typeName == p2.typeName
    struct_attrs = [s for s in dir(p1.value) if not s.startswith("_")]
    for atr in struct_attrs:
        assert getattr(p1, atr) == getattr(p2, atr)


def test_detect_interface(loader: XComponentLoader) -> None:
    assert loader is not None
    assert Info.is_interface_obj(loader)
