from pathlib import Path
import pytest
from unittest.mock import patch

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


import xml.etree.ElementTree as ET


def test_get_reg_item_prop_builtin_value(copy_fix_xml) -> None:
    xml_path: Path = copy_fix_xml("registymodifications.xml")

    item = "Writer/Layout/Other/TabStop"
    xpath = f'.//item[@oor:path="/org.openoffice.Office.{item}"]/value'

    tree = ET.parse(xml_path)
    root = tree.getroot()

    # element = root.find(xpath)
    element = tree.find(xpath, {"oor": "http://openoffice.org/2001/registry"})
    assert element is not None
    assert element.tag == "value"
    assert element.text == "1250"


def test_get_reg_item_prop_builtin_property_value(copy_fix_xml) -> None:
    xml_path: Path = copy_fix_xml("registymodifications.xml")

    item = "Calc/Calculate/Other"
    prop = "DecimalPlaces"
    xpath = f'.//item[@oor:path="/org.openoffice.Office.{item}"]/prop[@oor:name="{prop}"]/value'

    tree = ET.parse(xml_path)
    root = tree.getroot()

    # element = root.find(xpath)
    element = tree.find(xpath, {"oor": "http://openoffice.org/2001/registry"})
    assert element is not None
    assert element.tag == "value"
    assert element.text == "65535"


def test_get_reg_item_prop_builtin_node_prop_value(copy_fix_xml) -> None:
    xml_path: Path = copy_fix_xml("registymodifications.xml")

    item = "Logging/Settings"
    node = "org.openoffice.logging.sdbc.DriverManager"
    prop = "LogLevel"
    xpath = (
        f'.//item[@oor:path="/org.openoffice.Office.{item}"]/node[@oor:name="{node}"]/prop[@oor:name="{prop}"]/value'
    )

    tree = ET.parse(xml_path)
    root = tree.getroot()

    # element = root.find(xpath)
    element = tree.find(xpath, {"oor": "http://openoffice.org/2001/registry"})
    assert element is not None
    assert element.tag == "value"
    assert element.text == "2147483647"


def test_get_reg_item_prop_lo_value(copy_fix_xml, monkeypatch) -> None:
    from ooodev.utils.info import Info

    xml_path: Path = copy_fix_xml("registymodifications.xml")

    def mock_method():
        return xml_path

    monkeypatch.setattr(Info, "get_reg_mods_path", mock_method)

    item = "Writer/Layout/Other/TabStop"
    result = Info.get_reg_item_prop(item=item)
    assert result == "1250"


def test_get_reg_item_prop_lo_property_value(copy_fix_xml, monkeypatch) -> None:
    from ooodev.utils.info import Info

    xml_path: Path = copy_fix_xml("registymodifications.xml")

    def mock_method():
        return xml_path

    monkeypatch.setattr(Info, "get_reg_mods_path", mock_method)

    item = "Calc/Calculate/Other"
    prop = "DecimalPlaces"
    result = Info.get_reg_item_prop(item=item, prop=prop)
    assert result == "65535"


def test_get_reg_item_prop_lo_node_property_value(copy_fix_xml, monkeypatch) -> None:
    from ooodev.utils.info import Info

    xml_path: Path = copy_fix_xml("registymodifications.xml")

    def mock_method():
        return xml_path

    monkeypatch.setattr(Info, "get_reg_mods_path", mock_method)

    item = "Logging/Settings"
    node = "org.openoffice.logging.sdbc.DriverManager"
    prop = "LogLevel"
    result = Info.get_reg_item_prop(item=item, prop=prop, node=node)
    assert result == "2147483647"
