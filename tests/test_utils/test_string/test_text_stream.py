import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.string.str_list import StrList
from ooodev.utils.string.text_stream import TextStream


def test_get_text_stream_from_bytes(loader):
    stream = TextStream.get_text_input_stream_from_bytes(64, 9, 33, -7, -16, 27, -122, 110)
    assert stream.available() > 0


def test_text_stream_from_lst_and_back(loader):
    in_stream = TextStream.get_text_input_stream_from_str_list(StrList(["Hello", "World"]))
    str_lst = TextStream.get_str_list_from_text_input_stream(in_stream, ";")
    in_stream.set_encoding
    assert str_lst.to_string() == "Hello;World"
