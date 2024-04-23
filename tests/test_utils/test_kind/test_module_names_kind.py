import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.kind.module_names_kind import ModuleNamesKind


def test_new():
    x = ModuleNamesKind(ModuleNamesKind.BASIC_IDE)
    assert x == ModuleNamesKind.BASIC_IDE

    for item in ModuleNamesKind:
        assert item == ModuleNamesKind(item)
        assert item == ModuleNamesKind(item.value)
        assert item == ModuleNamesKind(item.value[0])
        assert item == ModuleNamesKind(item.value[1])
