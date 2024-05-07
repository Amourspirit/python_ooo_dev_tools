from pathlib import Path
import pytest
import uno
from ooo.dyn.xml.dom.node_type import NodeType
from ooodev.io.xml.xml import XML

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


def test_create_pay(loader, copy_fix_xml) -> None:
    xml_path: Path = copy_fix_xml("pay.xml")
    doc = XML.load_doc(xml_path)
    assert doc is not None


def test_get_xml_str(loader, copy_fix_xml) -> None:
    xml_path: Path = copy_fix_xml("pay.xml")
    doc = XML.load_doc(xml_path)
    assert doc is not None
    string = XML.get_xml_string(doc.getFirstChild())
    assert string is not None


def test_xml_string_to_doc(loader, copy_fix_xml) -> None:
    xml_path: Path = copy_fix_xml("pay.xml")
    with open(xml_path, "r") as f:
        xml_str = f.read()
    doc = XML.str_to_doc(xml_str)
    assert doc is not None


def test_xml_save_doc(loader, copy_fix_xml) -> None:
    xml_path: Path = copy_fix_xml("pay.xml")
    doc = XML.load_doc(xml_path)
    assert doc is not None
    xml_path = xml_path.with_name("pay2.xml")
    XML.save_doc(doc, xml_path)
    assert xml_path.exists()
    with open(xml_path, "r") as f:
        xml_str = f.read()
    assert xml_str is not None


def test_get_all_node_values(loader, copy_fix_xml) -> None:
    xml_path: Path = copy_fix_xml("pay.xml")
    doc = XML.load_doc(xml_path)
    assert doc is not None
    pays = doc.getElementsByTagName("payment")
    data = XML.get_all_node_values(pays, ("purpose", "amount", "tax", "maturity"))

    assert data is not None
    assert data[0] == ["Purpose", "Amount", "Tax", "Maturity"]
    assert data[1] == ["CD", "12.95", "19.1234", "2008-03-01"]
    assert data[2] == ["DVD", "19.95", "19.4321", "2008-03-02"]
    assert data[3] == ["Clothes", "99.95", "18.5678", "2008-03-03"]
    assert data[4] == ["Book", "9.49", "18.9876", "2008-03-04"]
