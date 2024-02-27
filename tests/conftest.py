from __future__ import annotations
import csv
import os
import sys
from pathlib import Path
import shutil
import stat
import tempfile
from typing import Any, Dict, List, TYPE_CHECKING, Optional
import pytest
from tests.fixtures import __test__path__ as test_fixture_path
from tests.fixtures.writer import __test__path__ as writer_fixture_path
from tests.fixtures.calc import __test__path__ as calc_fixture_path
from tests.fixtures.xml import __test__path__ as xml_fixture_path
from tests.fixtures.image import __test__path__ as img_fixture_path
from tests.fixtures.presentation import __test__path__ as pres_fixture_path
from ooodev.loader import Lo
from ooodev.utils import paths as mPaths
from ooodev.loader.inst import Options as LoOptions
from ooodev.conn import connectors

# from ooodev.connect import connectors as mConnectors
from ooodev.conn import cache as mCache

if TYPE_CHECKING:
    from com.sun.star.frame import XComponentLoader
else:
    XComponentLoader = object

# Snap Testing
# Limited Snap testing can be done.
# Mostly it is limited because the snap can't access the real tmp directory.
# To test snap the following must be modified:
# 1. soffice_path()
# 2. soffice_env()
# 3. loader()
# see the comments in each

# os.environ["ODEV_TEST_CONN_SOCKET"] = "true"
# os.environ["ODEV_TEST_HEADLESS"] = "0"
# os.environ[
#     "ODEV_CONN_SOFFICE"
# ] = "D:\\Portables\\PortableApps\\LibreOfficePortable\App\\libreoffice\\program\\soffice.exe"
# NOTE: No success running a portable version on windows from virtual environment.


# region deferred import testing
# https://docs.pytest.org/en/7.2.x/example/parametrize.html#deferring-the-setup-of-parametrized-resources


# def pytest_generate_tests(metafunc):
#     if "db" in metafunc.fixturenames:
#         metafunc.parametrize("importer", ["ooodev"], indirect=True)


# class IMPORT1:
#     "one database object"
#     pass


# @pytest.fixture
# def importer(request):
#     if request.param == "ooodev":
#         return IMPORT1()
#     # elif request.param == "d2":
#     #     return DB2()
#     else:
#         raise ValueError("importer - invalid internal test config")


# endregion deferred import testing


def remove_readonly(func, path, excinfo):
    try:
        os.chmod(path, stat.S_IWRITE)
        func(path)
    except Exception:
        pass


@pytest.fixture(scope="session")
def tmp_path_session():
    result = Path(tempfile.mkdtemp())  # type: ignore
    yield result
    if os.path.exists(result):
        shutil.rmtree(result, onerror=remove_readonly)


@pytest.fixture(scope="session")
def run_headless():
    # windows/powershell
    #   $env:NO_HEADLESS='1'; pytest; Remove-Item Env:\NO_HEADLESS
    # linux
    #  NO_HEADLESS="1" pytest
    return os.environ.get("ODEV_TEST_HEADLESS", "1") == "1"


@pytest.fixture(scope="session")
def fix_printer_name():
    # a printer name that is available on the test system
    # such as "Brother MFC-L2750DW series".
    # Printers such as "Microsoft Print to PDF" will not work for test.
    # see test_calc/test_calc_print.py
    return ""


@pytest.fixture(autouse=True)
def skip_for_headless(request, run_headless: bool):
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
                pytest.skip("Skipped in headless mode")


@pytest.fixture(autouse=True)
def skip_not_headless_os(request, run_headless: bool):
    # Usage:
    # @pytest.mark.skip_not_headless_os("linux", "Errors When GUI is present")
    # def test_write(loader, para_text) -> None:

    if not run_headless:
        rq = request.node.get_closest_marker("skip_not_headless_os")
        if rq:
            is_os = sys.platform.startswith(rq.args[0])
            if not is_os:
                return
            reason = ""
            try:
                reason = rq.args[1]
            except Exception:
                pass
            if reason:
                pytest.skip(reason)
            else:
                pytest.skip(f"Skipped in GUI mode on os: {rq.args[0]}")


@pytest.fixture(scope="session")
def soffice_path():
    # allow for a little more development flexibility
    # it is also fine to return "" or None from this function

    # return Path("/snap/bin/libreoffice")

    return mPaths.get_soffice_path()


@pytest.fixture(scope="session")
def soffice_env():
    # for snap testing the PYTHONPATH must be set to the virtual environment
    return {}
    # py_pth = mPaths.get_virtual_env_site_packages_path()
    # py_pth += f":{Path.cwd()}"
    # return {"PYTHONPATH": py_pth}


# region Loader methods
def _get_loader_pipe_default(
    headless: bool, soffice: str, working_dir: Any, env_vars: Optional[Dict[str, str]] = None
) -> XComponentLoader:
    dynamic = os.environ.get("ODEV_TEST_OPT_DYNAMIC", "") == "1"
    verbose = os.environ.get("ODEV_TEST_OPT_VERBOSE", "1") == "1"
    visible = os.environ.get("ODEV_TEST_OPT_VISIBLE", "") == "1"
    return Lo.load_office(
        connector=connectors.ConnectPipe(headless=headless, soffice=soffice, env_vars=env_vars, invisible=not visible),
        cache_obj=mCache.Cache(working_dir=working_dir),
        opt=LoOptions(verbose=verbose, dynamic=dynamic),
    )


def _get_loader_socket_default(
    headless: bool, soffice: str, working_dir: Any, env_vars: Optional[Dict[str, str]] = None
) -> XComponentLoader:
    dynamic = os.environ.get("ODEV_TEST_OPT_DYNAMIC", "") == "1"
    host = os.environ.get("ODEV_TEST_CONN_SOCKET_HOST", "localhost")
    port = int(os.environ.get("ODEV_TEST_CONN_SOCKET_PORT", 2002))
    verbose = os.environ.get("ODEV_TEST_OPT_VERBOSE", "1") == "1"
    visible = os.environ.get("ODEV_TEST_OPT_VISIBLE", "") == "1"
    return Lo.load_office(
        connector=connectors.ConnectSocket(
            host=host, port=port, headless=headless, soffice=soffice, env_vars=env_vars, invisible=not visible
        ),
        cache_obj=mCache.Cache(working_dir=working_dir),
        opt=LoOptions(verbose=verbose, dynamic=dynamic),
    )


def _get_loader_socket_no_start(
    headless: bool, working_dir: Any, env_vars: Optional[Dict[str, str]] = None
) -> XComponentLoader:
    dynamic = os.environ.get("ODEV_TEST_OPT_DYNAMIC", "") == "1"
    host = os.environ.get("ODEV_TEST_CONN_SOCKET_HOST", "localhost")
    port = int(os.environ.get("ODEV_TEST_CONN_SOCKET_PORT", 2002))
    verbose = os.environ.get("ODEV_TEST_OPT_VERBOSE", "1") == "1"
    visible = os.environ.get("ODEV_TEST_OPT_VISIBLE", "") == "1"
    return Lo.load_office(
        connector=connectors.ConnectSocket(
            host=host, port=port, headless=headless, start_office=False, env_vars=env_vars, invisible=not visible
        ),
        cache_obj=mCache.Cache(working_dir=working_dir),
        opt=LoOptions(verbose=verbose, dynamic=dynamic),
    )


@pytest.fixture(scope="session")
def loader(tmp_path_session, run_headless, soffice_path, soffice_env):
    # for testing with a snap the cache_obj must be omitted.
    # This because the snap is not allowed to write to the real tmp directory.
    test_socket = os.environ.get("ODEV_TEST_CONN_SOCKET", "0")
    connect_kind = os.environ.get("ODEV_TEST_CONN_SOCKET_KIND", "default")
    if test_socket == "1":
        if connect_kind == "no_start":
            loader = _get_loader_socket_no_start(
                headless=run_headless, working_dir=tmp_path_session, env_vars=soffice_env
            )
        else:
            loader = _get_loader_socket_default(
                headless=run_headless, soffice=soffice_path, working_dir=tmp_path_session, env_vars=soffice_env
            )
    else:
        loader = _get_loader_pipe_default(
            headless=run_headless, soffice=soffice_path, working_dir=tmp_path_session, env_vars=soffice_env
        )
    yield loader
    if connect_kind == "no_start":
        # only close office if it was started by the test
        return
    Lo.close_office()


# endregion Loader methods


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
    Converts a property string such as produced by
    Props.show_obj_props() into a key value pair.

    First line is expected to be Property Name and
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


@pytest.fixture(scope="session")
def formula_text() -> str:
    return "{{{sqrt{4x}} over 5} + {8 over 2}={4 over 3}}"
