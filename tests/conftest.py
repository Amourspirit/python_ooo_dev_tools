# coding: utf-8
import os
from pathlib import Path
import shutil
import tempfile
import pytest
from tests.fixtures import __test__path__ as fixture_path
from tests.fixtures.writer import __test__path__ as writer_fixture_path
from tests.fixtures.calc import __test__path__ as calc_fixture_path
from tests.fixtures.xml import __test__path__ as xml_fixture_path
from tests.fixtures.image import __test__path__ as img_fixture_path
from ooodev.utils.lo import Lo


@pytest.fixture(scope="session")
def loader():
    loader = Lo.load_office(host="localhost", port=2002)
    yield loader
    Lo.close_office()


@pytest.fixture(scope="function")
def tmp_path_fn():
    result = Path(tempfile.mkdtemp())
    yield result
    if os.path.exists(result):
        shutil.rmtree(result)


@pytest.fixture(scope="session")
def tmp_path():
    result = Path(tempfile.mkdtemp())
    yield result
    if os.path.exists(result):
        shutil.rmtree(result)


@pytest.fixture(scope="session")
def copy_fix_writer(tmp_path):
    def copy_res(fnm):
        src = Path(writer_fixture_path, fnm)
        dst = Path(tmp_path, fnm)
        shutil.copy2(src=src, dst=dst)
        return dst

    return copy_res


@pytest.fixture(scope="session")
def fix_writer_path():
    def get_res(fnm):
        return Path(writer_fixture_path, fnm)

    return get_res


@pytest.fixture(scope="session")
def copy_fix_calc(tmp_path):
    def copy_res(fnm):
        src = Path(calc_fixture_path, fnm)
        dst = Path(tmp_path, fnm)
        shutil.copy2(src=src, dst=dst)
        return dst

    return copy_res


@pytest.fixture(scope="session")
def fix_calc_path():
    def get_res(fnm):
        return Path(calc_fixture_path, fnm)

    return get_res

@pytest.fixture(scope="session")
def copy_fix_xml(tmp_path):
    def copy_res(fnm):
        src = Path(xml_fixture_path, fnm)
        dst = Path(tmp_path, fnm)
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
        def get_kv(kv_str:str) -> tuple:
            parts = kv_str.split(sep=":", maxsplit=1)
            if len(parts) == 0:
                return parts[0].strip(), ""
            return parts[0].strip(), parts[1].strip()
            
        lines = props_str.splitlines()
        dlines = {}
        frst = lines.pop(0) # first line is name such as Cursor Properties
        dlines["TEST_NAME"] = frst.strip()
        for line in lines:
            try:
                k, v = get_kv(line)
                dlines[k] = v
            except Exception:
                pass
        return dlines
    return text_to_dict