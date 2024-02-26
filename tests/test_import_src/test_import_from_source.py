import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])
# from tests.test_import_src._import_from_source import _test_imports_from_source, get_modules
from ._import_from_source import _test_imports_from_source, get_modules

# tests.test_calc.test_build_table
# add_module_organization_tests(__name__)
# add_module_organization_tests2("ooodev")


def test_dummy() -> None:
    print("Dummy test")
    # add_module_organization_tests("ooodev")
    assert True


@pytest.mark.parametrize("module", get_modules("ooodev"))
def test_eval(module: str):
    _test_imports_from_source(module)


def test_eval2():
    _test_imports_from_source("ooodev.write.style.write_style_family")
