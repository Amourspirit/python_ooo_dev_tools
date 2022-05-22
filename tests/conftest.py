# coding: utf-8
import os
from pathlib import Path
import shutil
import tempfile
import pytest
from tests.fixtures import __test__path__ as fixture_path
from tests.fixtures.writer import __test__path__ as writer_fixture_path
from ooodev.utils.lo import Lo

@pytest.fixture(scope='session')
def loader():
    loader = Lo.load_office(using_pipes=False)
    yield loader
    Lo.close_office()


@pytest.fixture(scope='function')
def tmp_path_fn():
    result = Path(tempfile.mkdtemp())
    yield result
    if os.path.exists(result):
        shutil.rmtree(result)

@pytest.fixture(scope='session')
def tmp_path():
    result = Path(tempfile.mkdtemp())
    yield result
    if os.path.exists(result):
        shutil.rmtree(result)

@pytest.fixture(scope='session')
def copy_fix_writer(tmp_path):
    def copy_res(doc_name: str):
        src = Path(writer_fixture_path, doc_name)
        dst = Path(tmp_path, doc_name)
        shutil.copy2(src=src, dst=dst)
        return dst
        
    return copy_res

@pytest.fixture(scope='session')
def fix_writer_path():
    def get_res(doc_name: str):
        return Path(writer_fixture_path, doc_name)
    return get_res