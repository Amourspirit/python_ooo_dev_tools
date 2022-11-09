from enum import IntFlag
from ooodev.utils.kind import kind_helper as Kh


class FlagKind(IntFlag):
    TEST_1 = 1
    TEST_2 = 2
    TEST_3 = 4
    TEST_4 = 8
    TEST_5 = TEST_3 | TEST_4


def test_from_str_int_enum() -> None:
    from ooodev.utils.kind.presentation_layout_kind import PresentationLayoutKind

    e1 = PresentationLayoutKind.CENTERED_TEXT
    e2 = Kh.enum_from_string(str(e1.value), PresentationLayoutKind)
    assert e1 == e2

    e2 = Kh.enum_from_string(f"0x{e1.value:02x}", PresentationLayoutKind)
    assert e1 == e2

    e2 = Kh.enum_from_string("CENTERED_TEXT", PresentationLayoutKind)
    assert e1 == e2

    e2 = Kh.enum_from_string("CENTERED TEXT", PresentationLayoutKind)
    assert e1 == e2

    e2 = Kh.enum_from_string("CENtERED-tEXT", PresentationLayoutKind)
    assert e1 == e2


def test_from_str_int_flag_enum() -> None:
    e1 = FlagKind.TEST_2
    e2 = Kh.enum_from_string(str(e1.value), FlagKind)
    assert e1 == e2

    e1 = FlagKind.TEST_5
    e2 = Kh.enum_from_string(str(e1.value), FlagKind)
    assert e1 == e2

    e2 = Kh.enum_from_string("TEST_5", FlagKind)
    assert e1 == e2

    e2 = Kh.enum_from_string("TEST 5", FlagKind)
    assert e1 == e2

    e2 = Kh.enum_from_string("tesT-5", FlagKind)
    assert e1 == e2

    e1 = FlagKind.TEST_2 | FlagKind.TEST_3
    e2 = Kh.enum_from_string(str(e1.value), FlagKind)
    assert e1 == e2

    e2 = Kh.enum_from_string(f"0x{e1.value:02x}", FlagKind)
    assert e1 == e2
