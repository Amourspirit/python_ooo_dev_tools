import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


# from com.sun.star.awt.tree import XTreeEditListener
# from com.sun.star.awt import XComboBox


@pytest.mark.parametrize(
    "import_str, alias",
    [
        pytest.param("com.sun.star.awt.XComboBox", "", id="XComboBox"),
        pytest.param("com.sun.star.awt.tree.XTreeEditListener", "", id="XTreeEditListener"),
        pytest.param(
            "ooodev.adapter.container.index_access_implement.IndexAccessImplement", "", id="IndexAccessImplement"
        ),
        pytest.param(
            "ooodev.adapter.container.index_access_implement.IndexAccessImplement", "", id="IndexAccessImplement2"
        ),
    ],
    ids=str,
)
def test_dynamic_from(import_str, alias, loader):
    from ooodev.utils.builder.dynamic_importer import DynamicImporter

    im = DynamicImporter.from_import(import_str, alias)
    if im.is_module_existing_import():
        im.unload_module()
    result = im.load_module()
    assert result is not None
    mt = result[0]
    clz = result[1]
    assert mt is not None
    assert clz is not None
    if not im.is_uno_import:
        assert im.is_module_existing_import()


@pytest.mark.parametrize(
    "import_str, alias",
    [
        pytest.param("verr", "Version", id="verr"),
    ],
    ids=str,
)
def test_direct_from(import_str, alias, loader):
    from ooodev.utils.builder.dynamic_importer import DynamicImporter

    im = DynamicImporter.from_direct_import(import_str, alias)
    if im.is_module_existing_import():
        im.unload_module()
    result = im.load_module()
    assert result is not None
    mt = result[0]
    assert mt is not None
    if not im.is_uno_import:
        assert im.is_module_existing_import()
