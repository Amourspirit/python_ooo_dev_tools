from multiprocessing import allow_connection_pickling
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
    result = gen_util.Util.is_iterable(arg, excluded_types)  # type: ignore
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


@pytest.mark.parametrize(
    "count,idx,allow_greater,result",
    (
        (1, 0, False, 0),
        (2, 0, False, 0),
        (3, -1, False, 2),
        (23, -1, False, 22),
        (23, -2, False, 21),
        (3, 1, False, 1),
        (10, -5, False, 5),
        (10, -2, False, 8),
        (10, -2, True, 8),
        (100, -3, False, 97),
        (3, 5, True, 3),
        (23, -1, True, 23),
        (23, -2, True, 21),
        (23, 23, True, 23),
        (23, 25, True, 23),
        (3, 0, True, 0),
        (3, 3, True, 3),
        (0, 0, True, 0),
        (0, -1, True, 0),
        (1, -1, True, 1),
        (2, -1, True, 2),
        (3, -1, True, 3),
    ),
)
def test_get_index(count: int, idx: int, allow_greater: bool, result: int) -> None:
    # count = 23, index is 22 idx = -2, allow_greater = False, result = 20

    # - 1 is last index, when count is 23 then last index is 22
    # - 2 is second last index, when count is 23 then second last index is 21
    # - 3 is third last index, when count is 23 then third last index is 20
    # - 4 is fourth last index, when count is 23 then fourth last index is 19

    # -1 is last index, when count is 10 then last index is 9
    # -2 is second last index, when count is 10 then second last index is 8
    # -3 is third last index, when count is 10 then third last index is 7
    # -4 is fourth last index, when count is 10 then fourth last index is 6
    # -5 is fifth last index, when count is 10 then fifth last index is 5
    # This means the math is count - abs(idx)
    index = gen_util.Util.get_index(idx=idx, count=count, allow_greater=allow_greater)
    assert result == index


@pytest.mark.parametrize(
    "count,idx,allow_greater",
    (
        (1, 1, False),
        (2, 2, False),
        (0, 0, False),
        (0, -1, False),
        (0, -2, False),
        (1, -2, False),
        (1, -2, True),
    ),
)
def test_get_index_error(count: int, idx: int, allow_greater: bool) -> None:
    with pytest.raises(IndexError):
        _ = gen_util.Util.get_index(idx=idx, count=count, allow_greater=allow_greater)
