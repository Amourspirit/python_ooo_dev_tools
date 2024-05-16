import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.string.str_list import StrList


def test_str_list() -> None:
    lst = ["A", "B", "C-With context", "D", "E", "F", "G", "H", "I", "J"]
    str_list = StrList(lst)
    for i, itm in enumerate(str_list):
        assert itm == lst[i]

    reversed_lst = lst[::-1]
    for i, itm in enumerate(reversed(str_list)):
        assert itm == reversed_lst[i]

    itm = str_list[1]
    assert itm == lst[1]

    lst_str = f"{str_list.separator}".join(lst)
    assert lst_str == str(str_list)

    slice1 = str_list[::-1]
    for i, itm in enumerate(slice1):
        assert itm == reversed_lst[i]

    slice2 = str_list[1:5]
    for i, itm in enumerate(slice2):
        assert itm == lst[i + 1]

    sub_list = str_list[1:5]
    sub_str = f"{slice2.separator}".join(sub_list)
    assert sub_str == str(sub_list)
