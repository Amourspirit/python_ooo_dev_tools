import csv
import os
from pathlib import Path
import shutil
import stat
import tempfile
from typing import List
import pytest
from tests.fixtures import __test__path__ as test_fixture_path
from tests.fixtures.writer import __test__path__ as writer_fixture_path
from tests.fixtures.calc import __test__path__ as calc_fixture_path
from tests.fixtures.xml import __test__path__ as xml_fixture_path
from tests.fixtures.image import __test__path__ as img_fixture_path
from tests.fixtures.presentation import __test__path__ as pres_fixture_path
from ooodev.utils.lo import Lo as mLo
from ooodev.utils import paths as mPaths

# from ooodev.connect import connectors as mConnectors
from ooodev.conn import cache as mCache

os.environ["NO_HEADLESS"] = "1"


def remove_readonly(func, path, excinfo):
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception:
        pass


@pytest.fixture(scope="session")
def tmp_path_session():
    result = Path(tempfile.mkdtemp())
    yield result
    if os.path.exists(result):
        shutil.rmtree(result, onerror=remove_readonly)


@pytest.fixture(scope="session")
def run_headless():
    # windows/powershell
    #   $env:NO_HEADLESS='1'; pytest; Remove-Item Env:\NO_HEADLESS
    # linux
    #  NO_HEADLESS="1" pytest
    no_headless = os.environ.get("NO_HEADLESS", "")
    if no_headless == "1":
        return False
    return True


@pytest.fixture(autouse=True)
def skip_for_headless(request, run_headless):
    # https://stackoverflow.com/questions/28179026/how-to-skip-a-pytest-using-an-external-fixture
    #
    # Also Added:
    # [tool.pytest.ini_options]
    # markers = ["skip_headless: skips a test in headless mode",]
    # see: https://docs.pytest.org/en/stable/how-to/mark.html
    #
    # Usage:
    # @pytest.mark.skip_headless("Requires Dispatch")
    # def test_write(loader, para_text) -> None:
    if run_headless:
        if request.node.get_closest_marker("skip_headless"):
            reason = ""
            try:
                reason = request.node.get_closest_marker("skip_headless").args[0]
            except Exception:
                pass
            if reason:
                pytest.skip(reason)
            else:
                pytest.skip("Skiped in headless mode")


@pytest.fixture(scope="session")
def soffice_path():
    # allow for a little more development flexibility
    # it is alos fine to return "" or None from this function
    return mPaths.get_soffice_path()


@pytest.fixture(scope="session")
def loader(tmp_path_session, run_headless, soffice_path):
    loader = mLo.load_office(
        connector=mLo.ConnectPipe(headless=run_headless, soffice=soffice_path),
        cache_obj=mCache.Cache(working_dir=tmp_path_session),
        opt=mLo.Options(verbose=True),
    )
    # loader = mLo.load_office(connector=mLo.ConnectSocket(headless=True, soffice=soffice_path), cache_obj=mCache.Cache(working_dir=tmp_path_session))
    yield loader
    mLo.close_office()


@pytest.fixture(scope="function")
def tmp_path_fn():
    result = Path(tempfile.mkdtemp())
    yield result
    if os.path.exists(result):
        shutil.rmtree(result, onerror=remove_readonly)


@pytest.fixture(scope="session")
def copy_fix_writer(tmp_path_session):
    def copy_res(fnm):
        src = Path(writer_fixture_path, fnm)
        dst = Path(tmp_path_session, fnm)
        shutil.copy2(src=src, dst=dst)
        return dst

    return copy_res


@pytest.fixture(scope="session")
def copy_fix_presentation(tmp_path_session):
    def copy_res(fnm):
        src = Path(pres_fixture_path, fnm)
        dst = Path(tmp_path_session, fnm)
        shutil.copy2(src=src, dst=dst)
        return dst

    return copy_res


@pytest.fixture(scope="session")
def fixture_path():
    return Path(test_fixture_path)


@pytest.fixture(scope="session")
def fix_writer_path():
    def get_res(fnm):
        return Path(writer_fixture_path, fnm)

    return get_res


@pytest.fixture(scope="session")
def copy_fix_calc(tmp_path_session):
    def copy_res(fnm):
        src = Path(calc_fixture_path, fnm)
        dst = Path(tmp_path_session, fnm)
        shutil.copy2(src=src, dst=dst)
        return dst

    return copy_res


@pytest.fixture(scope="session")
def fix_calc_path():
    def get_res(fnm):
        return Path(calc_fixture_path, fnm)

    return get_res


@pytest.fixture(scope="session")
def copy_fix_xml(tmp_path_session):
    def copy_res(fnm):
        src = Path(xml_fixture_path, fnm)
        dst = Path(tmp_path_session, fnm)
        shutil.copy2(src=src, dst=dst)
        return dst

    return copy_res


@pytest.fixture(scope="session")
def fix_xml_path():
    def get_res(fnm: str):
        return Path(xml_fixture_path, fnm)

    return get_res


@pytest.fixture(scope="session")
def fix_image_path():
    def get_res(fnm):
        return Path(img_fixture_path, fnm)

    return get_res


@pytest.fixture(scope="session")
def props_str_to_dict():
    """
    Convets a property string such as produced by
    Props.show_obj_props() into a key value pair.

    First line is excpected to be Property Name and
    is added as Key TEST_NAME

    Returns:
        Returns a dictionary of Key value strings
    """

    def text_to_dict(props_str: str) -> dict:
        def get_kv(kv_str: str) -> tuple:
            parts = kv_str.split(sep=":", maxsplit=1)
            if len(parts) == 0:
                return parts[0].strip(), ""
            return parts[0].strip(), parts[1].strip()

        lines = props_str.splitlines()
        dlines = {}
        frst = lines.pop(0)  # first line is name such as Cursor Properties
        dlines["TEST_NAME"] = frst.strip()
        for line in lines:
            try:
                k, v = get_kv(line)
                dlines[k] = v
            except Exception:
                pass
        return dlines

    return text_to_dict


@pytest.fixture(scope="session")
def bond_movies_lst_dict(fix_writer_path) -> List[dict]:
    csv_file = Path(writer_fixture_path, "bondMovies.txt")
    results = []
    with open(csv_file, "r", newline="") as csv_file:
        csv_reader = csv.reader(csv_file, delimiter="\t")
        line_count = 0
        cols = []
        for row in csv_reader:
            if line_count in (0, 1, 2, 4):  # not csv lines
                line_count += 1
                continue
            if line_count == 3:  # column headers
                cols.extend(row)
            else:
                results.append(dict(zip(cols, row)))
            line_count += 1
    return results


@pytest.fixture(scope="session")
def bond_movies_table(fix_writer_path) -> List[list]:
    csv_file = Path(writer_fixture_path, "bondMovies.txt")
    results = []
    with open(csv_file, "r", newline="") as csv_file:
        csv_reader = csv.reader(csv_file, delimiter="\t")
        line_count = 0
        for row in csv_reader:
            if line_count in (0, 1, 2, 4):  # not csv lines
                line_count += 1
                continue
            # first row will be column names
            results.append(row)
            line_count += 1
    return results


@pytest.fixture(scope="session")
def para_text() -> str:
    p_txt = (
        "To Sherlock Holmes she is always THE woman. I have seldom heard"
        " him mention her under any other name. In his eyes she eclipses"
        " and predominates the whole of her sex. It was not that he felt"
        " any emotion akin to love for Irene Adler. All emotions, and that"
        " one particularly, were abhorrent to his cold, precise but"
        " admirably balanced mind. He was, I take it, the most perfect"
        " reasoning and observing machine that the world has seen, but as a"
        " lover he would have placed himself in a false position. He never"
        " spoke of the softer passions, save with a gibe and a sneer. They"
        " were admirable things for the observer--excellent for drawing the"
        " veil from men's motives and actions. But for the trained reasoner"
        " to admit such intrusions into his own delicate and finely"
        " adjusted temperament was to introduce a distracting factor which"
        " might throw a doubt upon all his mental results. Grit in a"
        " sensitive instrument, or a crack in one of his own high-power"
        " lenses, would not be more disturbing than a strong emotion in a"
        " nature such as his. And yet there was but one woman to him, and"
        " that woman was the late Irene Adler, of dubious and questionable memory."
    )
    return p_txt
