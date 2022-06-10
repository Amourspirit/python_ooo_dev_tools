from pathlib import Path
import pytest
from typing import cast

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.xml_util import XML

_PAY_INDENTED = """<?xml version="1.0" ?>
<payments>
	<payment>
		<purpose>CD</purpose>
		<amount>12.95</amount>
		<tax>19.1234</tax>
		<maturity>2008-03-01</maturity>
	</payment>
	<payment>
		<purpose>DVD</purpose>
		<amount>19.95</amount>
		<tax>19.4321</tax>
		<maturity>2008-03-02</maturity>
	</payment>
	<payment>
		<purpose>Clothes</purpose>
		<amount>99.95</amount>
		<tax>18.5678</tax>
		<maturity>2008-03-03</maturity>
	</payment>
	<payment>
		<purpose>Book</purpose>
		<amount>9.49</amount>
		<tax>18.9876</tax>
		<maturity>2008-03-04</maturity>
	</payment>
</payments>
"""

# region    Sheet Methods
def test_indent_pay_doc(fix_xml_path) -> None:
    xml_path: Path = fix_xml_path("pay.xml")
    with open(xml_path, "r") as xfile:
        xstr = xfile.read()
    # remove all whitespace and line breaks
    xml_str = "".join([s.strip() for s in xstr.splitlines()])

    xdoc = XML.str_to_doc(xml_str=xml_str)
    indent_str = XML.indent(xdoc)
    assert indent_str == _PAY_INDENTED


def test_indent_pay_str(fix_xml_path) -> None:
    xml_path: Path = fix_xml_path("pay.xml")
    with open(xml_path, "r") as xfile:
        xstr = xfile.read()
    # remove all whitespace and line breaks
    xml_str = "".join([s.strip() for s in xstr.splitlines()])

    indent_str = XML.indent(xml_str)
    assert indent_str == _PAY_INDENTED


def test_indent_pay_path(fix_xml_path) -> None:
    xml_path: Path = fix_xml_path("pay.xml")
    indent_str = XML.indent(xml_path)
    assert indent_str == _PAY_INDENTED


def test_indent_pay_bad_type() -> None:
    with pytest.raises(TypeError) as err_info:
        XML.indent(2)
    e = err_info.value
    assert cast(str, e.args[0]).endswith("Got int")
