from pathlib import Path
import pytest
# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


# region    Sheet Methods
def test_create_pay(copy_fix_xml, capsys: pytest.CaptureFixture) -> None:
    xml_path: Path = copy_fix_xml("pay.xml")
    from ooodev.utils.lo import Lo
    from ooodev.utils.xml_util import XML
    xdoc = XML.load_doc(str(xml_path))
    pays = xdoc.getElementsByTagName("payment")
    assert pays is not None
    data = XML.get_all_node_values(pays, ("purpose", "amount", "tax", "maturity"))
    capsys.readouterr() # clear buffer
    Lo.print_table(name="payments", table=data)
    cap_result = capsys.readouterr()
    cap_out = cap_result.out
    assert cap_out == """-- payments ----------------
CD  12.95  19.1234  2008-03-01
DVD  19.95  19.4321  2008-03-02
Clothes  99.95  18.5678  2008-03-03
Book  9.49  18.9876  2008-03-04

"""