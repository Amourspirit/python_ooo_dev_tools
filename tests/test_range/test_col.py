from typing import cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_cell_obj_to_cell_values() -> None:
    from ooodev.utils.data_type.col_obj import ColObj

    c1 = ColObj.from_int(1)
    assert c1.index == 0
    assert c1.value == "A"

    c1 = ColObj.from_int(0, True)
    assert c1.index == 0
    assert c1.value == "A"

    c1 = ColObj("A")
    assert c1.index == 0
    assert c1.value == "A"
    assert c1 == "A"

    c2 = ColObj("b")
    assert c2.index == 1
    assert c2.value == "B"
    assert c2 == "b"

    assert c2 > c1
    assert c1 < c2

    assert c2 == "B"
    assert "B" == c2

    assert c1 == "a"
    assert "a" == c1

    assert c2 > "a"
    assert "A" < c2

    assert c2 >= "b"
    assert c2 >= "a"

    assert "c" > c2
    assert c2 < "c"

    c3 = cast(ColObj, c2 + 1)
    assert c3.index == 2
    assert c3.value == "C"
    assert c3 == "C"

    c5 = ColObj("E")
    assert c5.index == 4

    c10 = cast(ColObj, c1 + 9)
    assert c10.value == "J"
    assert c10.index == 9

    c11 = cast(ColObj, c2 + 9)
    assert c11.value == "K"
    assert c11.index == 10
    assert c11 > c10

    # I is adding 9
    # b (2) + I (9) = K (11)
    c11 = cast(ColObj, c2 + "i")
    assert c11.value == "K"
    assert c11.index == 10
    assert c11 > c10

    c11 = cast(ColObj, "i" + c2)
    assert c11.value == "K"
    assert c11.index == 10
    assert c11 > c10
    assert c11 < "O"

    c9 = cast(ColObj, c11 - 2)
    assert c9.value == "I"
    assert c9.index == 8
    assert c9 < c11
    assert c9 < "K"

    c9 = c11 - c2
    assert c9.value == "I"
    assert c9.index == 8
    assert c9 < c11
    assert c9 < "K"

    c9 = cast(ColObj, c11 - "B")
    assert c9.value == "I"
    assert c9.index == 8
    assert c9 < c11
    assert c9 < "K"

    c2 = cast(ColObj, 5 - c3)
    assert c2.index == 1

    c2 = cast(ColObj, "E" - c3)
    assert c2.index == 1

    assert c1.index == 0
    assert c2.index == 1
    assert c3.index == 2
    cc = c1 + c2
    assert cc.index == 2
    assert c1 + c2 == c3

    c_sum = sum([c1, c2, c3])
    assert c_sum.index == 5

    c_sum = cast(ColObj, sum([c1, 2, 3]))
    assert c_sum.index == 5

    c_sum = cast(ColObj, sum([c1, "b", "c"]))
    assert c_sum.index == 5

    c_sum = cast(ColObj, sum([c1, "b", 3]))
    assert c_sum.index == 5
