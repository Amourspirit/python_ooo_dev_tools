import pytest
from ooodev.utils import color


@pytest.mark.parametrize(
    "c_val",
    [
        "3",
        "4",
        "0x800080",
        "0xFF1493",
        "FF1493",
        "78849553",
        "0X800080",
        "DIM_GRAY",
        "DIM GRAY",
        "DIM-GRAY",
        "Dim_Gray",
        "LIGHT_SEA_GREEN",
        "LIGHT_SEA GREEN",
        "LIGHT SEA-GREEN",
        "light sea green",
        "MEDIUM_PURPLE",
        "ORANGE_RED",
        "MEDIUM VIOLET RED",
        "hot_pink",
    ],
)
def test_from_str(c_val: str) -> None:
    c = color.CommonColor.from_str(c_val)
    assert c >= 0


@pytest.mark.parametrize(
    "c_val",
    [
        "nan",
        "not a color",
        "light  sea  green",
        "DUDE_NOT_HERE",
    ],
)
def test_from_str_err(c_val: str) -> None:
    with pytest.raises(ValueError):
        color.CommonColor.from_str(c_val)
