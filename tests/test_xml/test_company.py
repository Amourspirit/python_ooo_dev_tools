from pathlib import Path
import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


def test_create_pay(copy_fix_xml) -> None:
    xml_path: Path = copy_fix_xml("company.xml")
    from ooodev.utils.xml_util import XML

    xdoc = XML.load_doc(xml_path)
    root = xdoc.childNodes
    assert root is not None
    comps = XML.get_node(tag_name="Companies", nodes=root)
    assert comps is not None
    comp = XML.get_node("Company", comps.childNodes)
    exec = XML.get_node("Executive", comp.childNodes)
    assert exec is not None

    exec_type = XML.get_node_attr("type", exec)
    assert exec_type == "CEO"
    ex_nodes = exec.childNodes
    last_name = XML.get_node_value("LastName", ex_nodes)
    assert last_name == "Smith"
    first_name = XML.get_node_value("FirstName", ex_nodes)
    assert first_name == "Jim"
    street = XML.get_node_value("street", ex_nodes)
    assert street == "123 Broad Street"
    city = XML.get_node_value("city", ex_nodes)
    assert city == "Manchester"
    state = XML.get_node_value("state", ex_nodes)
    assert state == "Cheshire"
    zip = XML.get_node_value("zip", ex_nodes)
    assert zip == "11234"

    comp = comp.nextSibling
    exec = XML.get_node("Executive", comp.childNodes)
    exec_type = XML.get_node_attr("type", exec)
    assert exec_type == "President"
    ex_nodes = exec.childNodes
    last_name = XML.get_node_value("LastName", ex_nodes)
    assert last_name == "Jones"
    first_name = XML.get_node_value("FirstName", ex_nodes)
    assert first_name == "Lucy"
    street = XML.get_node_value("street", ex_nodes)
    assert street == "23 Bradford St"
    city = XML.get_node_value("city", ex_nodes)
    assert city == "Asbury"
    state = XML.get_node_value("state", ex_nodes)
    assert state == "Lincs"
    zip = XML.get_node_value("zip", ex_nodes)
    assert zip == "33451"

    comp = comp.nextSibling
    exec = XML.get_node("Executive", comp.childNodes)
    exec_type = XML.get_node_attr("type", exec)
    assert exec_type == "Boss"
    ex_nodes = exec.childNodes
    last_name = XML.get_node_value("LastName", ex_nodes)
    assert last_name == "Singh"
    first_name = XML.get_node_value("FirstName", ex_nodes)
    assert first_name == "Oxley"
    street = XML.get_node_value("street", ex_nodes)
    assert street == "16d Towers"
    city = XML.get_node_value("city", ex_nodes)
    assert city == "Wimbledon"
    state = XML.get_node_value("state", ex_nodes)
    assert state == "London"
    zip = XML.get_node_value("zip", ex_nodes)
    assert zip == "77392"

    # get all the data in the tree for a given node/tag name
    lnames = ("Smith", "Jones", "Singh")
    ln_nodes = xdoc.getElementsByTagName("LastName")
    assert len(ln_nodes) == 3
    for ln in ln_nodes:
        name = XML.get_node_value(ln)
        assert name in lnames
