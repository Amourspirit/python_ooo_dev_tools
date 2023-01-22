from __future__ import annotations
from ooodev.utils.unit_convert import UnitConvert, Length
import pytest
import math


@pytest.mark.parametrize(
    ("val", "expected"),
    [
        (100, 2540.0),
        (100.3, 2547.62),
        (53, 1346.2),
        (0.5, 12.7),
        (0.2, 5.08),
    ],
)
def test_convert_inch_mm(val: int | float, expected: int) -> None:
    result = UnitConvert.convert(val, Length.IN, Length.MM)
    assert math.isclose(result, expected)


@pytest.mark.parametrize(
    ("val", "expected"),
    [
        (1, 0.3527777777778),
        (7, 2.469444444444),
        (53, 18.69722222222),
        (71, 25.04722222222),
        (176, 62.08888888889),
        (35, 12.347222222229),
        (318, 112.1833333333),
    ],
)
def test_convert_pt_mm(val: int | float, expected: int) -> None:
    result = UnitConvert.convert(val, Length.PT, Length.MM)
    assert math.isclose(result, expected)


@pytest.mark.parametrize(
    ("val", "expected"),
    [
        (0.1, 3.527777777778),
        (0.2, 7.055555555556),
        (0.5, 17.63888888889),
        (0.7, 24.69444444444),
        (0.75, 26.45833333333),
        (1, 35.27777777778),
        (2, 70.55555555556),
        (7, 246.9444444444),
        (53, 1869.722222222),
        (71, 2504.722222222),
        (176, 6208.888888889),
        (35, 1234.7222222229),
        (318, 11218.33333333),
    ],
)
def test_convert_pt_mm100(val: int | float, expected: int) -> None:
    result = UnitConvert.convert(val, Length.PT, Length.MM100)
    assert math.isclose(result, expected)


@pytest.mark.parametrize(
    ("val", "expected"),
    [
        (0.5, 10.0),
        (0.6, 12.0),
        (0.7, 14.0),
        (0.75, 15.0),
        (1, 20.0),
        (2, 40.0),
    ],
)
def test_convert_pt_twip(val: int | float, expected: int) -> None:
    result = UnitConvert.to_twips(val, Length.PT)
    assert math.isclose(result, expected)


@pytest.mark.parametrize(
    ("val", "expected"),
    [
        (1, 1.763888888889),
        (10, 17.63888888889),
        (15, 26.45833333333),
        (18, 31.75),
        (20, 35.27777777778),
        (21, 37.04166666667),
        (25, 44.09722222222),
        (26, 45.86111111111),
        (2243, 3956.402777778),
    ],
)
def test_convert_twip_mm100(val: int | float, expected: int) -> None:
    result = UnitConvert.convert_twip_mm100(val)
    assert math.isclose(result, expected)


@pytest.mark.parametrize(
    ("val", "expected"),
    [
        (4, 2.267716535433),
        (2000, 1133.858267717),
        (1.763888888889, 1.0),
        (17.63888888889, 10.0),
        (26.45833333333, 15.0),
        (31.75, 18.0),
        (35.27777777778, 20.0),
        (37.04166666667, 21.0),
        (44.09722222222, 25.0),
        (45.86111111111, 26.0),
        (3956.402777778, 2243.0),
    ],
)
def test_convert_mm100_twip(val: int | float, expected: int) -> None:
    result = UnitConvert.convert(val, Length.MM100, Length.TWIP)
    assert math.isclose(result, expected)


@pytest.mark.parametrize(
    ("val", "expected"),
    [(2, 609.6), (2.9, 883.92), (14, 4267.2)],
)
def test_convert_ft_mm(val: int | float, expected: int) -> None:
    result = UnitConvert.convert(val, Length.FT, Length.MM)
    assert math.isclose(result, expected)
