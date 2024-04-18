import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])
# from tests.test_import_src._import_from_source import _test_imports_from_source, get_modules
from ._import_from_source import _test_imports_from_source, get_modules


@pytest.mark.parametrize("module", get_modules("ooodev"))
def test_ooodev_modules(module):
    _test_imports_from_source(module)


def test_ooodev_imports_single():
    # this method can be use for specific debugging.
    # _test_imports_from_source("ooodev.adapter.configuration.configuration_access_comp")
    _test_imports_from_source("")
