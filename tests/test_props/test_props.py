import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


# region    Sheet Methods
def test_show_indexed_props(loader, capsys: pytest.CaptureFixture) -> None:
    from ooodev.utils.props import Props
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    try:
        sheets = Calc.get_sheets(doc)
        capsys.readouterr() # clear buffer
        Props.show_indexed_props("Sheets Props", sheets)
        cap_result = capsys.readouterr()
        cap_out: str = cap_result.out
        assert isinstance(cap_out, str)
        cap_lst = [line.strip() for line in cap_out.splitlines()]
        assert len(cap_lst) > 0
        assert cap_lst[3].startswith('AbsoluteName: $Sheet1.$A$1')
    finally:
        Lo.close_doc(doc=doc, deliver_ownership=False)

def test_prop_value_to_string(loader) -> None:
    from ooodev.utils.props import Props
    from ooodev.utils.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.utils.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    try:
        sheet = Calc.get_sheet(doc=doc, index=0)
        prop_str = Props.prop_value_to_string(sheet)
        prop_lst = [line.strip() for line in prop_str.splitlines()]
        assert len(prop_lst) > 0
        assert prop_lst[1].startswith('AbsoluteName = $Sheet1.$A$1')
    finally:
        Lo.close_doc(doc=doc, deliver_ownership=False)