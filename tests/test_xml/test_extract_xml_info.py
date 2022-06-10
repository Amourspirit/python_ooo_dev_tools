from pathlib import Path
import pytest
import io
import shutil
import xml.dom.minidom as md
# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


from ooodev.utils.xml_util import XML

# region    Sheet Methods
def test_extract_pay(fix_xml_path, copy_fix_xml, tmp_path) -> None:
    xml_path: Path = copy_fix_xml("pay.xml")
    xdoc = XML.load_doc(str(xml_path))
    out_fnm = "pay.xml.txt"
    out_path = Path(tmp_path, out_fnm)
    sw = io.StringIO()
    root = xdoc.firstChild
    for child in root.childNodes:
        visit_node(pw=sw,node=child, ind="")
        sw.write("\n")

    with open(out_path, 'w') as fd:
        sw.seek(0)
        shutil.copyfileobj(sw, fd)
    
    assert out_path.exists()
    with open(out_path, 'r') as new_file:
        out_txt = new_file.read()
    assert out_txt is not None
    with open(out_path, 'r') as new_file:
        with open(fix_xml_path(out_fnm), 'r') as existing:
            assert [row for row in new_file] == [row for row in existing]
    assert out_txt == \
"""payment  purpose: "CD"  amount: "12.95"  tax: "19.1234"  maturity: "2008-03-01"
payment  purpose: "DVD"  amount: "19.95"  tax: "19.4321"  maturity: "2008-03-02"
payment  purpose: "Clothes"  amount: "99.95"  tax: "18.5678"  maturity: "2008-03-03"
payment  purpose: "Book"  amount: "9.49"  tax: "18.9876"  maturity: "2008-03-04"
"""

def test_extract_clubs(fix_xml_path, copy_fix_xml, tmp_path) -> None:
    xml_path: Path = copy_fix_xml("clubs.xml")
    xdoc = XML.load_doc(str(xml_path))
    out_fnm = "clubs.xml.txt"
    out_path = Path(tmp_path, out_fnm)
    sw = io.StringIO()
    root = xdoc.firstChild
    for child in root.childNodes:
        visit_node(pw=sw,node=child, ind="")
        sw.write("\n")

    with open(out_path, 'w') as fd:
        sw.seek(0)
        shutil.copyfileobj(sw, fd)
    
    assert out_path.exists()

    with open(out_path, 'r') as new_file:
        with open(fix_xml_path(out_fnm), 'r') as existing:
            assert [row for row in new_file] == [row for row in existing]

def test_extract_weather(fix_xml_path, copy_fix_xml, tmp_path) -> None:
    xml_path: Path = copy_fix_xml("weather.xml")
    xdoc = XML.load_doc(str(xml_path))
    out_fnm = "weather.xml.txt"
    out_path = Path(tmp_path, out_fnm)
    sw = io.StringIO()
    root = xdoc.firstChild
    for child in root.childNodes:
        visit_node(pw=sw,node=child, ind="")
        sw.write("\n")

    with open(out_path, 'w') as fd:
        sw.seek(0)
        shutil.copyfileobj(sw, fd)
    
    assert out_path.exists()
    with open(out_path, 'r') as new_file:
        with open(fix_xml_path(out_fnm), 'r') as existing:
            assert [row for row in new_file] == [row for row in existing]

def test_extract_company(fix_xml_path, copy_fix_xml, tmp_path) -> None:
    xml_path: Path = copy_fix_xml("company.xml")
    xdoc = XML.load_doc(str(xml_path))
    out_fnm = "company.xml.txt"
    out_path = Path(tmp_path, out_fnm)
    sw = io.StringIO()
    root = xdoc.firstChild
    for child in root.childNodes:
        visit_node(pw=sw,node=child, ind="")
        sw.write("\n")

    with open(out_path, 'w') as fd:
        sw.seek(0)
        shutil.copyfileobj(sw, fd)
    
    assert out_path.exists()
    with open(out_path, 'r') as new_file:
        with open(fix_xml_path(out_fnm), 'r') as existing:
            assert [row for row in new_file] == [row for row in existing]


def visit_node(pw:io.StringIO, node: md.Node, ind: str):
    """Visit a node by printing its name, any attribute data,
     any text node data, and then recursively visiting 
     all the node's children."""
    pw.write(f"{ind}{node.nodeName}")
    visit_attrs(pw, node)
    
    # examine all the child nodes
    for child in node.childNodes:
        if child.nodeType == md.Node.TEXT_NODE:
            trimmed_val = child.data.strip()
            if len(trimmed_val) == 0:
                 pw.write("\n")
            else:
                 pw.write(f': "{trimmed_val}"')
            # element names with values end with ':'
        elif child.nodeType == md.Node.ELEMENT_NODE:
            visit_node(pw, child, ind + "  ")
            
   

def visit_attrs(pw:io.StringIO, node: md.Node) -> None:
    """print all the attributes -- name and data"""
    # attrs is {} if there are no attributes
    # attrs: dict = node.attributes
    if node.attributes is None:
        return
    attrs = dict(node.attributes.items())
    for k, v in attrs.items():
         pw.write(f'  {k}= "{v}"')
    # attribute names end with '='
