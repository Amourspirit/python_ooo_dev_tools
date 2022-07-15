import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])



def test_doc_info(loader, copy_fix_presentation) -> None:
    
    from ooodev.utils.lo import Lo
    from ooodev.utils.info import Info
    test_doc = copy_fix_presentation("algs.odp")
    doc = Lo.open_doc(fnm=test_doc, loader=loader)
    Info.print_doc_properties(doc)
    