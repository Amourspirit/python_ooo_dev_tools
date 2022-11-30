import sys
from pathlib import Path
from ooodev.utils.file_io import FileIO
import pytest


@pytest.mark.parametrize(
    "input,dir,",
    [
        ("program", "program"),
        ("ab3/123", "ab3/123"),
        ("rgb/myfile.txt", "rgb"),
        ("pics/png/myfile.jpg", "pics/png/"),
        ("", ""),
    ],
)
def test_make_dir(input, dir, tmp_path) -> None:
    p = Path(tmp_path, input)
    dp = Path(tmp_path, dir)
    result = FileIO.make_directory(p)
    assert result == dp
    assert result.exists()


@pytest.mark.skipif(sys.platform == "win32", reason="does not run on windows")
@pytest.mark.parametrize(
    "input, expected",
    [
        ("file:///home/user/myfile.txt", "/home/user/myfile.txt"),
        ("file:///var/log/myfile.txt", "/var/log/myfile.txt"),
    ],
)
def test_linux_url_to_path(input: str, expected: str) -> None:
    result = FileIO.url_to_path(input)
    ep = Path(expected)
    assert result == ep


@pytest.mark.skipif(sys.platform != "win32", reason="only runs on windows")
@pytest.mark.parametrize(
    "input, expected",
    [
        ("file:///c:/user/myfile.txt", "c:/user/myfile.txt"),
        ("file:///d:/mypath/myfile.txt", "d:/mypath/myfile.txt"),
        ("file:///d:/myfile.txt", "d:/myfile.txt"),
    ],
)
def test_win_url_to_path(input: str, expected: str) -> None:
    # result = FileIO.url_to_path(input)
    result = FileIO.uri_to_path(input)
    ep = Path(expected)
    assert result == ep

    result = FileIO.url_to_path(input)
    assert result == ep


@pytest.mark.skipif(sys.platform != "win32", reason="only runs on windows")
@pytest.mark.parametrize(
    "input, expected",
    [("c:/user/myfile.txt", "file:///c:/user/myfile.txt"), ("d:/myfile.txt", "file:///d:/myfile.txt")],
)
def test_win_path_to_url(input: str, expected: str) -> None:
    result = FileIO.fnm_to_url(input)
    assert result.casefold() == expected.casefold()


@pytest.mark.skipif(sys.platform == "win32", reason="does not run on windows")
@pytest.mark.parametrize(
    "input, expected",
    [("/home/user/myfile.txt", "file:///home/user/myfile.txt"), ("/home/myfile.txt", "file:///home/myfile.txt")],
)
def test_win_path_to_url(input: str, expected: str) -> None:
    result = FileIO.fnm_to_url(input)
    assert result.casefold() == expected.casefold()


@pytest.mark.parametrize(
    "input, expected",
    [("/home/user/myfile.txt", "myfile.txt"), ("/home/myPic.jpeg", "myPic.jpeg"), ("", "")],
)
def test_get_fnm(input: str, expected: str) -> None:
    result = FileIO.get_fnm(input)
    assert result == expected


@pytest.mark.parametrize(
    "input",
    ["myfile.txt", "myPic.jpeg"],
)
def test_is_openable(input, tmp_path) -> None:
    p = Path(tmp_path, input)
    p.touch()  # create fiele
    result = FileIO.is_openable(p)
    assert result


@pytest.mark.parametrize(
    "input",
    ["myfile.txt", "myPic.jpeg"],
)
def test_not_is_openable(input, tmp_path) -> None:
    p = Path(tmp_path, input)
    result = FileIO.is_openable(p)
    assert result == False


@pytest.mark.parametrize(
    "input,expected",
    [("myfile.txt", True), ("myPic.jpeg", True), (Path("somepath"), True), ("", False), (None, False)],
)
def test_is_valid_path_or_str(input, expected: bool) -> None:
    result = FileIO.is_valid_path_or_str(input)
    assert result == expected


def test_is_exist_file(tmp_path) -> None:
    p = Path(tmp_path, "myfile.txt")
    p.touch()
    result = FileIO.is_exist_file(p)
    assert result


def test_is_exist_file_dir(tmp_path) -> None:
    result = FileIO.is_exist_file(tmp_path)
    assert result == False


def test_is_exist_file_invalid_path() -> None:
    with pytest.raises(ValueError):
        result = FileIO.is_exist_file("", True)


def test_is_exist_file_not_file(tmp_path) -> None:
    with pytest.raises(ValueError):
        result = FileIO.is_exist_file(tmp_path, True)


def test_is_exist_file_not_exist(tmp_path) -> None:
    p = Path(tmp_path, "random.txt")
    with pytest.raises(FileNotFoundError):
        result = FileIO.is_exist_file(p, True)


def test_is_exist_dir(tmp_path) -> None:
    result = FileIO.is_exist_dir(tmp_path)
    assert result


def test_is_exist_dir_not_exist(tmp_path) -> None:
    result = FileIO.is_exist_dir(Path(tmp_path, "nope"))
    assert result == False


def test_is_exist_dir_invalid_path() -> None:
    with pytest.raises(ValueError):
        result = FileIO.is_exist_dir("", True)


def test_is_exist_dir_not_exist(tmp_path) -> None:
    p = Path(tmp_path, "nope")
    with pytest.raises(FileNotFoundError):
        result = FileIO.is_exist_dir(p, True)


def test_create_and_del_tmp_file() -> None:
    tmp_file = FileIO.create_temp_file("jpeg")
    p = Path(tmp_file)
    assert p.exists()
    assert p.is_file()
    FileIO.delete_file(p)
    assert p.exists() == False


def test_save_str() -> None:
    tmp_file = FileIO.create_temp_file("txt")
    s = "hello world"
    p = Path(tmp_file)
    assert p.exists()

    FileIO.save_string(tmp_file, s)
    with open(p, "r") as file:
        f_txt = file.read()
    FileIO.delete_file(tmp_file)
    assert f_txt == s


def test_save_bytes() -> None:
    tmp_file = FileIO.create_temp_file("bin")
    p = Path(tmp_file)
    assert p.exists()
    b = b"hello world"

    FileIO.save_bytes(tmp_file, b)
    with open(p, "rb") as file:
        b_val = file.read()
    FileIO.delete_file(tmp_file)
    assert b_val == b


def test_save_array(bond_movies_table) -> None:
    tmp_file = FileIO.create_temp_file("txt")
    try:
        FileIO.save_array(tmp_file, bond_movies_table)
        with open(tmp_file, "r") as file:
            txt = file.read()
        assert txt.startswith("Title\t")
    finally:
        FileIO.delete_file(tmp_file)


def test_append_to() -> None:
    tmp_file = FileIO.create_temp_file("txt")
    try:
        s = "Hello World"
        p = Path(tmp_file)
        assert p.exists()

        FileIO.save_string(p, s)
        FileIO.append_to(p, "\nNice Day")
        with open(p, "r") as file:
            f_txt = file.read()
        assert f_txt == "Hello World\nNice Day\n"
    finally:
        FileIO.delete_file(tmp_file)


def test_uri_absolute() -> None:
    input = "file:///C:/Program%20Files/LibreOffice/program/../share/gallery/sounds/apert2.wav"
    expected = "file:///C:/Program%20Files/LibreOffice/share/gallery/sounds/apert2.wav"
    result = FileIO.uri_absolute(input)
    assert result == expected
