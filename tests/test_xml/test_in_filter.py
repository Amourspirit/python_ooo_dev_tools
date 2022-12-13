"""
Important: XML.apply_xslt() is not availabe in macros.
     XML.apply_xslt() requires lxml python package

Convert XML to an Office document in two steps:
     1) use the supplied XSLT to convert the XML
        into Flat XML understood by Office;

     2) Use the correct Flat XML import filter to load
        the flat XML data into Office
"""
import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.xml_util import XML
from ooodev.utils.lo import Lo
from ooodev.utils.file_io import FileIO
from ooodev.utils.info import Info
from ooodev.utils.gui import GUI
from ooodev.office.calc import Calc
from ooodev.office.write import Write

# region    Sheet Methods
def test_transform_pay(loader, copy_fix_xml) -> None:
    """
    - Copy pay.xml and payImport.xsl to tmp dir
    - transform pay.xml using payImport.xls and store in xml_str
    - save xml_str as xml file in tmp directory
    - create an empty file in temp dir with ods extension
    - Open transformed xml in calc as a spreadsheet
    - Save calc spreadsheet
    - close calc
    - Open that saved spreadsheet in calc.
    - get the expected range and test
    """
    visible = False
    delay = 0  # 1000
    pay_import = copy_fix_xml("payImport.xsl")
    pay = copy_fix_xml("pay.xml")

    xml_str = XML.apply_xslt(xml_fnm=str(pay), xls_fnm=pay_import)
    assert xml_str is not None

    # save flat XML data to temp file
    flat_fnm = FileIO.create_temp_file("xml")
    FileIO.save_string(flat_fnm, xml_str)

    # create a Calc File
    ods_fnm = FileIO.create_temp_file("ods")

    # open temp file using Office's correct Flat XML filter
    doc_type = Lo.ext_to_doc_type(Info.get_ext(ods_fnm))
    assert doc_type == Lo.DocTypeStr.CALC
    doc = Lo.open_flat_doc(fnm=flat_fnm, doc_type=doc_type, loader=loader)
    assert doc is not None
    GUI.set_visible(is_visible=visible, odoc=doc)

    Lo.delay(delay)
    Lo.save_doc(doc=doc, fnm=ods_fnm)
    Lo.close_doc(doc=doc)
    Lo.delay(1000)

    doc = Calc.open_doc(fnm=ods_fnm, loader=loader)
    sheet = Calc.get_sheet(doc=doc, index=0)

    arr = Calc.get_array(sheet=sheet, range_name="A1:D5")
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    Lo.delay(delay)
    # (('Purpose', 'Amount', 'Tax', 'Maturity'), ('CD', 12.95, 19.1234, 39508.0), ('DVD', 19.95, 19.4321, 39509.0), ('Clothes', 99.95, 18.5678, 39510.0), ('Book', 9.49, 18.9876, 39511.0))
    assert arr[0][0] == "Purpose"
    assert arr[0][1] == "Amount"
    assert arr[0][2] == "Tax"
    assert arr[0][3] == "Maturity"

    assert arr[1][0] == "CD"
    assert arr[1][3] == 39508.0

    assert arr[2][0] == "DVD"
    assert arr[2][3] == 39509.0

    assert arr[3][0] == "Clothes"
    assert arr[3][3] == 39510.0

    assert arr[4][0] == "Book"
    assert arr[4][3] == 39511.0

    Lo.close(closeable=doc, deliver_ownership=False)
    Lo.delay(1000)


def test_transform_clubs(loader, copy_fix_xml) -> None:
    """
    - Copy clubs.xml and clubsImport.xsl to tmp dir
    - transform clubs.xml using clubsImport.xls and store in xml_str
    - save xml_str as xml file in tmp directory
    - create an empty file in temp dir with odt extension
    - Open transformed xml in Write
    - Save Write doc
    - close Write
    - Open saved Write doc
    - get lines and test
    """
    # for unknown reason gettin a strange fail on Ubuntu 20.04. Not tested on other os at this time.
    # tests/test_xml/test_in_filter.py .terminate called after throwing an instance of 'com::sun::star::lang::DisposedException'
    # If this test is run via VS Code plugin it passes. If run command line via pytest tests/ then if conditionally fails.
    # the conditions are as follows:
    #   Is not giving error in Calc
    #   Is only erroring if window visibality is false. Setting visibility after document is loaded seems ok.
    #   Is only erroring when opening a flat doc ( in Write ).
    #
    # Solution:
    # The current working solution is to not have flat docs open hidden. This was the case in the original java code.
    # changes in Lo.open_flat_doc()
    #       Changed
    #       return cls.open_doc(fnm, loader, mProps.Props.make_props(FilterName=nn, Hidden=False))
    #       To
    #       return cls.open_doc(fnm, loader, mProps.Props.make_props(FilterName=nn))
    visible = False
    delay = 0  # 1000
    clubs_import = copy_fix_xml("clubsImport.xsl")
    clubs = copy_fix_xml("clubs.xml")

    xml_str = XML.apply_xslt(xml_fnm=clubs, xls_fnm=str(clubs_import))
    assert xml_str is not None

    # save flat XML data to temp file
    flat_fnm = FileIO.create_temp_file("xml")
    FileIO.save_string(flat_fnm, xml_str)

    # create a Calc File
    odt_fnm = FileIO.create_temp_file("odt")

    # open temp file using Office's correct Flat XML filter
    doc_type = Lo.ext_to_doc_type(Info.get_ext(odt_fnm))
    assert doc_type == Lo.DocTypeStr.WRITER
    doc = Lo.open_flat_doc(fnm=flat_fnm, doc_type=doc_type, loader=loader)
    assert doc is not None
    GUI.set_visible(is_visible=visible, odoc=doc)

    Lo.delay(delay)
    Lo.save_doc(doc=doc, fnm=odt_fnm)
    Lo.close_doc(doc=doc)
    Lo.delay(1000)

    doc = Write.open_doc(fnm=odt_fnm, loader=loader)

    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    Lo.delay(delay)

    cursor = Write.get_cursor(doc)
    text = Write.get_all_text(cursor=cursor)
    lines = text.splitlines()
    lines_len = len(lines)
    assert lines_len == 1967
    assert lines[0] == "BAWA"
    assert lines[41] == "Godfather Wrestling Club H14"
    assert lines[86] == "Email: titanwrestling@example.com"
    assert lines[734] == "Chartered: 2002"
    assert lines[890] == "Phone: 707-555-5970"
    assert lines[993] == "Contact: Duane Fidel"
    assert lines[1361] == "Contact: Randy Campbell"
    assert lines[1928] == "Nipomo Youth Wrestling Club I28"

    Lo.close(closeable=doc, deliver_ownership=True)
    Lo.delay(1000)


# endregion    Sheet Methods
