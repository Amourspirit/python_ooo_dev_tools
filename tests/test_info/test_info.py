import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


def test_get_doc_type(loader, fix_writer_path) -> None:

    from ooodev.utils.lo import Lo
    from ooodev.utils.info import Info

    test_doc = fix_writer_path("story.odt")
    type = Info.get_doc_type(test_doc)
    assert type == "writer8"


def test_get_available_services(loader, copy_fix_writer) -> None:

    from ooodev.utils.lo import Lo
    from ooodev.utils.info import Info

    test_doc = copy_fix_writer("story.odt")
    doc = Lo.open_doc(fnm=test_doc, loader=loader)
    try:
        services = Info.get_available_services(doc)
        assert len(services) > 10  # 208 on Windows 10, LO 7.3
    finally:
        Lo.close_doc(doc=doc, deliver_ownership=False)


def test_get_interfaces(loader, copy_fix_writer) -> None:

    from ooodev.utils.lo import Lo
    from ooodev.utils.info import Info

    test_doc = copy_fix_writer("story.odt")
    doc = Lo.open_doc(fnm=test_doc, loader=loader)
    try:
        interfaces = Info.get_interfaces(doc)
        assert len(interfaces) > 10  # 72 on Windows 10, LO 7.3
    finally:
        Lo.close_doc(doc=doc, deliver_ownership=False)


def test_get_methods(loader) -> None:
    from ooodev.utils.info import Info

    methods = Info.get_methods("com.sun.star.text.XTextDocument")
    assert len(methods) > 10 # 19 on Windows 10, LO 7.3

def test_get_methods_obj(loader) -> None:

    from ooodev.utils.lo import Lo
    from ooodev.utils.info import Info
    from ooodev.office.write import Write
    from ooo.dyn.beans.property_concept import PropertyConceptEnum

    doc = Write.create_doc(loader=loader)
    try:
        methods_all = Info.get_methods_obj(obj=doc)
        assert len(methods_all) > 10 # 199 on Windows 10, LO 7.3
        properties = Info.get_methods_obj(obj=doc, property_concept=PropertyConceptEnum.PROPERTYSET)
        assert len(properties) > 10 # 68 on Windows 10, LO 7.3
        attrs = Info.get_methods_obj(obj=doc, property_concept=PropertyConceptEnum.ATTRIBUTES)
        assert len(attrs) > 10 # 34 on Windows 10, LO 7.3
        methods = Info.get_methods_obj(obj=doc, property_concept=PropertyConceptEnum.METHODS)
        assert len(methods) == 0
    finally:
        Lo.close_doc(doc=doc, deliver_ownership=False)