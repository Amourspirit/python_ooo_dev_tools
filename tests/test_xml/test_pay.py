from pathlib import Path
import pytest
import uno
from ooo.dyn.xml.dom.node_type import NodeType

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


def test_create_pay(loader, copy_fix_xml) -> None:
    from ooodev.adapter.xml.dom.document_builder_comp import DocumentBuilderComp
    from ooodev.adapter.xml.xpath.x_path_api_comp import XPathAPIComp

    builder = DocumentBuilderComp.from_lo()

    xml_path: Path = copy_fix_xml("pay.xml")
    uri = uno.systemPathToFileUrl(str(xml_path))

    doc = builder.parse_uri(uri)
    assert doc is not None
    doc.normalize()
    xpath = XPathAPIComp.from_lo()
    pays = xpath.select_node_list(doc.getFirstChild(), "payment")
    assert pays is not None
    assert len(pays) == 4
    node = pays[0]
    assert node is not None
    assert node.getNodeType() == NodeType.ELEMENT_NODE
