import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])


def test_show_indexed_props(loader, capsys: pytest.CaptureFixture) -> None:
    from ooodev.utils.props import Props
    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.gui.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    try:
        sheets = Calc.get_sheets(doc)
        capsys.readouterr()  # clear buffer
        Props.show_indexed_props("Sheets Props", sheets)
        cap_result = capsys.readouterr()
        cap_out: str = cap_result.out
        assert isinstance(cap_out, str)
        cap_lst = [line.strip() for line in cap_out.splitlines()]
        assert len(cap_lst) > 0
        assert cap_lst[3].startswith("AbsoluteName: $Sheet1.$A$1")
    finally:
        Lo.close_doc(doc=doc, deliver_ownership=False)


def test_prop_value_to_string(loader) -> None:
    from ooodev.utils.props import Props
    from ooodev.loader.lo import Lo
    from ooodev.office.calc import Calc

    # from ooodev.gui.gui import GUI

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    try:
        sheet = Calc.get_sheet(doc=doc, index=0)
        prop_str = Props.prop_value_to_string(sheet)
        prop_lst = [line.strip() for line in prop_str.splitlines()]
        assert len(prop_lst) > 0
        assert prop_lst[1].startswith("AbsoluteName = $Sheet1.$A$1")
    finally:
        Lo.close_doc(doc=doc, deliver_ownership=False)


def test_prop_url(loader, fix_writer_path):
    from ooodev.utils.props import Props
    from ooodev.loader.lo import Lo
    from com.sun.star.frame import XModel

    test_doc = fix_writer_path("scandalStart.odt")
    doc = Lo.open_doc(fnm=test_doc, loader=loader)
    try:
        model = Lo.qi(XModel, doc)
        args = model.getArgs()
        url = str(Props.get_value(name="URL", props=args))
        assert url.endswith("scandalStart.odt")
    finally:
        Lo.close_doc(doc, False)


def test_prop_get_with_default(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.utils.props import Props
    from ooodev.office.calc import Calc

    assert loader is not None
    doc = Calc.create_doc(loader)
    assert doc is not None
    try:
        sheet = Calc.get_sheet(doc=doc, index=0)
        idx = 3
        height = 14
        cell_range = Calc.set_row_height(sheet=sheet, height=height, idx=idx)
        assert cell_range is not None
        c_height = Props.get(cell_range, "Height")
        assert c_height is not None
        p = Props.get(cell_range, "non-exisgint-prop", None)
        assert p is None
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)
