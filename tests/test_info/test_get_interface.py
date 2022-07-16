import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])



def test_get_interfaces(loader, copy_fix_writer) -> None:
    
    from ooodev.utils.lo import Lo
    from ooodev.utils.info import Info
    test_doc = copy_fix_writer("story.odt")
    doc = Lo.open_doc(fnm=test_doc, loader=loader)
    try:
        interfaces = Info.get_interfaces(doc)
        assert len(interfaces) == 72
    finally:
        Lo.close_doc(doc, False)