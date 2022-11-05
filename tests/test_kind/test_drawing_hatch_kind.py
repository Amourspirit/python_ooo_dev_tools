import pytest
from ooodev.utils.kind.drawing_hatching_kind import DrawingHatchingKind


@pytest.mark.parametrize(
    "input,expected_name",
    (
        ("BLACK_0_DEGREES", "BLACK_0_DEGREES"),
        ("BLACK 0 DEGREES", "BLACK_0_DEGREES"),
        ("BLACK-0-DEGREES", "BLACK_0_DEGREES"),
        ("black-0-DEGREES", "BLACK_0_DEGREES"),
        ("black 0 DEGREES", "BLACK_0_DEGREES"),
        ("Black0Degrees", "BLACK_0_DEGREES"),
        ("Red45Degrees", "RED_45_DEGREES"),
    ),
)
def test_from_str(input: str, expected_name: str) -> None:
    result = DrawingHatchingKind.from_str(input)
    assert result.name == expected_name
