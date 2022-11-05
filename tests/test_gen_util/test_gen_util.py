import pytest
from typing import Iterable
from ooodev.utils import gen_util


@pytest.mark.parametrize(
    "arg,excluded_types,expected",
    (
        (("f", "f"), None, True),
        (["f", "f"], None, True),
        (iter("ff"), None, True),
        (range(44), None, True),
        (b"ff", None, True),
        ("ff", None, False),  # u"ff"
        ("ff", (), True),  # remove default ignore str
        (44, None, False),
        (gen_util.Util.to_camel_case, None, False),  #  function
        (["f", "f"], (list,), False),
    ),
)
def test_is_iterable(arg: object, excluded_types: Iterable[type], expected: bool) -> None:
    result = gen_util.Util.is_iterable(arg, excluded_types)
    assert result == expected


@pytest.mark.parametrize(
    "s,expected",
    (
        ("redGreenColor", "red_green_color"),
        ("red_green_color", "red_green_color"),
        ("RedGreenColor", "red_green_color"),
        ("redGreenColor", "red_green_color"),
        ("RedGreen_color", "red_green_color"),
        ("red_green_color", "red_green_color"),
        ("", ""),
        ("Red", "red"),
        ("Red0Green", "red_0_green"),
        ("2500RedGreen", "2500_red_green"),
        ("  2500RedGreen  ", "2500_red_green"),
        ("RedGreen44822 ", "red_green_44822"),
        ("Red0green", "red_0_green"),
        ("Red0green1", "red_0_green_1"),
        ("12Red0green1", "12_red_0_green_1"),
        ("18Red22green100", "18_red_22_green_100"),
    ),
)
def test_to_snake(s: str, expected: str) -> None:
    result = gen_util.Util.to_snake_case(s)
    assert result == expected


@pytest.mark.parametrize(
    "s,expected",
    (
        ("redGreenColor", "RED_GREEN_COLOR"),
        ("red_green_color", "RED_GREEN_COLOR"),
        ("RedGreenColor", "RED_GREEN_COLOR"),
        ("redGreenColor", "RED_GREEN_COLOR"),
        ("RedGreen_color", "RED_GREEN_COLOR"),
        ("red_green_color", "RED_GREEN_COLOR"),
        ("", ""),
        ("Red", "RED"),
        ("18Red22green100", "18_RED_22_GREEN_100"),
    ),
)
def test_to_snake_upper(s: str, expected: str) -> None:
    result = gen_util.Util.to_snake_case_upper(s)
    assert result == expected


@pytest.mark.parametrize(
    "s,expected",
    (
        ("red_green_color", "redGreenColor"),
        ("red_grEEn_color", "redGreenColor"),
        ("RedGreenColor", "redGreenColor"),
        ("Red", "red"),
        ("", ""),
        ("X", "x"),
        ("y", "y"),
    ),
)
def test_to_pascal(s: str, expected: str) -> None:
    result = gen_util.Util.to_pascal_case(s)
    assert result == expected


@pytest.mark.parametrize(
    "s,expected",
    (
        ("red_green_color", "RedGreenColor"),
        ("red_grEEn_color", "RedGreenColor"),
        ("Red", "Red"),
        ("red", "Red"),
        ("", ""),
        ("X", "X"),
        ("y", "Y"),
        ("pascalCase", "PascalCase"),
    ),
)
def test_to_camel(s: str, expected: str) -> None:
    result = gen_util.Util.to_camel_case(s)
    assert result == expected


@pytest.mark.parametrize(
    "s,strip,expected",
    (
        ("two  spaces", False, "two spaces"),
        ("red    grEEn  color", False, "red grEEn color"),
        ("  two  spaces  ", False, " two spaces "),
        ("  two  spaces  ", True, "two spaces"),
        ("  two  more\n  spaces  ", True, "two more\n spaces"),
    ),
)
def test_to_single_space(s: str, strip: bool, expected: str) -> None:
    result = gen_util.Util.to_single_space(s, strip)
    assert result == expected


@pytest.mark.parametrize(
    "s,expected",
    (
        ("  two  spaces  ", "two spaces"),
        ("  two  more\n  spaces  ", "two more\n spaces"),
    ),
)
def test_to_single_space_default(s: str, expected: str) -> None:
    result = gen_util.Util.to_single_space(s)
    assert result == expected
