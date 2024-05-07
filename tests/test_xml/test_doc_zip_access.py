import pytest
import uno
from ooodev.io.zip.zip import ZIP
from ooodev.write import WriteDoc

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


def test_get_styles(loader, copy_fix_writer) -> None:
    from ooodev.adapter.xml.dom.document_builder_comp import DocumentBuilderComp
    from ooodev.utils.string.text_stream import TextStream
    from ooodev.io.xml.xml import XML

    test_doc = copy_fix_writer("scandalStart.odt")
    doc = WriteDoc.open_doc(fnm=test_doc, loader=loader)
    # doc = WriteDoc.create_doc(loader=loader)
    # with doc.lo_inst.global_event_broadcaster:
    # suppress global document events.
    content_bytes = ZIP.get_zip_content_as_byte_array(doc.component, "content.xml")
    txt_comp = TextStream.get_text_input_stream_from_bytes(*content_bytes)
    ox = DocumentBuilderComp.from_lo()
    dom_content = ox.parse(txt_comp.component)
    automatic_styles = dom_content.getElementsByTagNameNS(
        "urn:oasis:names:tc:opendocument:xmlns:office:1.0", "automatic-styles"
    ).item(0)

    xml_str = XML.get_xml_string(automatic_styles)
    assert xml_str is not None
