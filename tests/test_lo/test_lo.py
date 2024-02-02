from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno


def test_bridge(loader) -> None:
    from ooodev.loader.lo import Lo

    assert Lo.bridge is not None
    assert Lo.is_loaded
    # don't know how to pass doc string to custom classproperty yet
    doc_str = Lo.bridge.__doc__
    assert True


def _test_dispose() -> None:
    # this test is to be run manually
    # if this test were to be run with other test
    # it would wipe out the connection to office because Lo is a static class.
    from ooodev.loader.lo import Lo

    with Lo.Loader(Lo.ConnectPipe(headless=True)) as loader:
        # confirm that Lo is connected to office
        assert Lo._xcc is not None
        assert Lo._mc_factory is not None
        assert Lo._lo_inst is not None
    # delay to ensure office bridge is disposed
    Lo.delay(1000)
    # now that bridge is disposed Lo internals should be reset.
    # this is because when Lo.load_office is called a listener is attached to
    # bridge via LoNamedEvent.OFFICE_LOADED event.
    assert Lo._xcc is None
    assert Lo._mc_factory is None
    assert Lo._lo_inst is None


def test_lo_inst(loader) -> None:
    from ooodev.loader.lo import Lo
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.loader.inst.doc_type import DocTypeStr
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI

    # version added: 0.9.8

    # using existing connection to office, create a new calc doc.
    # Get the sheet and set the value of the first cell.
    # create a new instance of LoInst.
    # Loads office by using the existing connection to office.
    # creates a new calc doc.
    # Get the sheet and set the value of the first cell.

    doc = Calc.create_doc()
    try:
        assert doc is not None
        if not Lo.bridge_connector.headless:
            GUI.set_visible(visible=True, doc=doc)
        sheet = Calc.get_sheet(doc, 0)
        Calc.set_val(value="LO TEST", sheet=sheet, cell_name="A1")
        assert sheet is not None

        lo = LoInst()
        lo.load_office(Lo.bridge_connector, opt=Lo.Options(verbose=True))
        lo_doc = lo.create_doc(DocTypeStr.CALC)
        assert lo_doc is not None
        try:
            if not Lo.bridge_connector.headless:
                GUI.set_visible(visible=True, doc=lo_doc)
            lo_sheet = Calc.get_sheet(lo_doc, 0)
            Calc.set_val(value="LO INST TEST", sheet=lo_sheet, cell_name="A1")
            assert lo_sheet is not None
        finally:
            lo.close_doc(lo_doc)
    finally:
        Lo.close_doc(doc)
