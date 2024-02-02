import pytest
from unittest.mock import patch
import uno


# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


def test_get_doc_type(loader, fix_writer_path) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.utils.info import Info

    test_doc = fix_writer_path("story.odt")
    type = Info.get_doc_type(test_doc)
    assert type == "writer8"


def test_get_available_services(loader, copy_fix_writer) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.utils.info import Info

    test_doc = copy_fix_writer("story.odt")
    doc = Lo.open_doc(fnm=test_doc, loader=loader)
    try:
        services = Info.get_available_services(doc)
        assert len(services) > 10  # 208 on Windows 10, LO 7.3
    finally:
        Lo.close_doc(doc=doc, deliver_ownership=False)


def test_get_interfaces(loader, copy_fix_writer) -> None:
    from ooodev.loader.lo import Lo
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
    assert len(methods) > 10  # 19 on Windows 10, LO 7.3


def test_get_methods_obj(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.utils.info import Info
    from ooodev.office.write import Write
    from ooo.dyn.beans.property_concept import PropertyConceptEnum

    doc = Write.create_doc(loader=loader)
    try:
        methods_all = Info.get_methods_obj(obj=doc)
        assert len(methods_all) > 10  # 199 on Windows 10, LO 7.3
        properties = Info.get_methods_obj(obj=doc, property_concept=PropertyConceptEnum.PROPERTYSET)
        assert len(properties) > 10  # 68 on Windows 10, LO 7.3
        attrs = Info.get_methods_obj(obj=doc, property_concept=PropertyConceptEnum.ATTRIBUTES)
        assert len(attrs) > 10  # 34 on Windows 10, LO 7.3
        methods = Info.get_methods_obj(obj=doc, property_concept=PropertyConceptEnum.METHODS)
        assert len(methods) == 0
    finally:
        Lo.close_doc(doc=doc, deliver_ownership=False)


def test_identify(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.utils.info import Info
    from ooodev.office.write import Write

    doc = Write.create_doc(loader=loader)
    try:
        identifier = Info.get_identifier(obj=doc)
        assert identifier == "com.sun.star.text.TextDocument"

        implementation = Info.get_implementation_name(obj=doc)
        assert implementation == "SwXTextDocument"
    finally:
        Lo.close_doc(doc=doc, deliver_ownership=False)


def test_gallery_dir(loader) -> None:
    # ensure Info.get_gallery_dir() is getting the correct location.
    from ooodev.utils.info import Info

    # Gallery file.
    fnm = Info.get_gallery_dir() / "sounds" / "applause.wav"
    assert fnm.exists()


def test_version_str(loader) -> None:
    from ooodev.utils.info import Info

    ver = Info.version
    assert isinstance(ver, str)


def test_version_info(loader) -> None:
    from ooodev.utils.info import Info

    ver = Info.version_info
    assert isinstance(ver, tuple)
    assert len(ver) >= 2
    assert isinstance(ver[0], int)
    assert isinstance(ver[1], int)


def test_info_theme(loader) -> None:
    from ooodev.utils.info import Info

    ver = Info.version_info
    theme = Info.get_office_theme()
    if ver >= (7, 4, 0, 0):
        assert len(theme) > 0
    assert isinstance(theme, str)


def test_info_language(loader) -> None:
    from ooodev.utils.info import Info

    lang = Info.language
    assert len(lang) > 4  # "en-US" or similar


def test_info_language_locale() -> None:
    from ooodev.utils.info import Info

    # mocker.patch("ooodev.utils.info.Info.get_config", return_value="en-US")
    # mocker.patch("ooodev.utils.info.Info._language", return_value="en-US")
    with patch("ooodev.utils.info.Info._language", "en-US", create=True):
        lang = Info.language_locale
        assert lang.Language == "en"
        assert lang.Country == "US"


# write a test for Info.parse_languange_code()
def test_parse_languange_code() -> None:
    from ooodev.utils.info import Info

    lang = Info.parse_language_code("en-US")
    assert lang.Language == "en"
    assert lang.Country == "US"
    assert lang.Variant == ""
    lang = Info.parse_language_code("en-GB")
    assert lang.Language == "en"
    assert lang.Country == "GB"
    assert lang.Variant == ""
    lang = Info.parse_language_code("en-gb")
    assert lang.Language == "en"
    assert lang.Country == "GB"
    assert lang.Variant == ""
    lang = Info.parse_language_code("en-GB-oxendict")
    assert lang.Language == "en"
    assert lang.Country == "GB"
    assert lang.Variant == "oxendict"
    lang = Info.parse_language_code("en--oxendict")
    assert lang.Language == "en"
    assert lang.Country == ""
    assert lang.Variant == "oxendict"
    lang = Info.parse_language_code("en-GB-oxendict-this-that")
    assert lang.Language == "en"
    assert lang.Country == "GB"
    assert lang.Variant == "oxendict-this-that"


def test_parese_language_code_error() -> None:
    from ooodev.utils.info import Info

    with pytest.raises(ValueError):
        _ = Info.parse_language_code("en_GB")
    with pytest.raises(ValueError):
        _ = Info.parse_language_code("-GB")
