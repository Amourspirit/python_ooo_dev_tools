import pytest
from ooodev.utils import color


@pytest.mark.parametrize(
    "input,expected",
    (
        ("3", 3),
        ("4", 4),
        ("0x800080", 0x800080),
        ("0xFF1493", 0xFF1493),
        ("FF1493", 0xFF1493),
        ("78849553", 78849553),
        ("0X800080", 0x800080),
        ("DIM_GRAY", 0x696969),
        ("DIM_GREY", 0x696969),
        ("DIM GRAY", 0x696969),
        ("DIM-GRAY", 0x696969),
        ("Dim_Gray", 0x696969),
        ("DimGray", 0x696969),
        ("dimGray", 0x696969),
        ("LIGHT_SEA_GREEN", 0x20B2AA),
        ("LIGHT_SEA GREEN", 0x20B2AA),
        ("LIGHT SEA-GREEN", 0x20B2AA),
        ("light sea green", 0x20B2AA),
        ("light   sea   green", 0x20B2AA),
        ("  light   sea   green  ", 0x20B2AA),
        ("LightSeaGreen", 0x20B2AA),
        ("lightSeaGreen", 0x20B2AA),
        ("MEDIUM_PURPLE", 0x9370DB),
        ("ORANGE_RED", 0xFF4500),
        ("MEDIUM VIOLET RED", 0xC71585),
        ("hot_pink", 0xFF69B4),
        ("red", 0xFF0000),
        ("blue", 0x0000FF),
        ("MEDIUM_SEA_GREEN", 0x3CB371),
        ("MEDIUM-SEA-GREEN", 0x3CB371),
        ("MEDIUM SEA GREEN", 0x3CB371),
        ("medium sea green", 0x3CB371),
        ("mediumSeaGreen", 0x3CB371),
        ("MediumSeaGreen", 0x3CB371),
    ),
)
def test_from_str(input: str, expected: int) -> None:
    c = color.CommonColor.from_str(input)
    assert c == expected


@pytest.mark.parametrize(
    "c_val",
    (
        "nan",
        "not a color",
        "light  sea  green",
        "DUDE_NOT_HERE",
    ),
)
def test_from_str_err(c_val: str) -> None:
    with pytest.raises(ValueError):
        color.CommonColor.from_str(c_val)
