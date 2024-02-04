import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


def test_service_names(loader, copy_fix_writer) -> None:

    from ooodev.loader.lo import Lo
    from ooodev.utils.info import Info

    test_doc = copy_fix_writer("story.odt")
    doc = Lo.open_doc(fnm=test_doc, loader=loader)
    try:
        names = Info.get_service_names(Lo.Service.WRITER)
        assert len(names) == 1

        names = Info.get_service_names()
        # on Window 10, LibreOffice 7.3.x. 1010 services reported
        assert len(names) > 200  # 1010
    finally:
        Lo.close_doc(doc, False)
