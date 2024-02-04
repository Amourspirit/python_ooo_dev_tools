import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


def test_print_doc_properties(loader, copy_fix_presentation) -> None:

    from ooodev.loader.lo import Lo
    from ooodev.utils.info import Info

    test_doc = copy_fix_presentation("algs.odp")
    doc = Lo.open_doc(fnm=test_doc, loader=loader)
    try:
        Info.print_doc_properties(doc)
    finally:
        Lo.close_doc(doc, False)
