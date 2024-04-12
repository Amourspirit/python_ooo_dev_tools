from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.gui.menu.shortcuts import Shortcuts
from ooodev.loader.inst.service import Service


def test_shortcut_global(loader) -> None:
    sc = Shortcuts()
    sc_all = sc.get_all()
    # get a list of shortcuts such as:
    # [('shift+ctrl+N', '.uno:NewDoc'), ('ctrl+O', '.uno:Open'), ('ctrl+S', '.uno:Save'), ('ctrl+V', '.uno:Paste'), ('ctrl+X', '.uno:Cut'), ...]
    assert sc_all
    cmd = sc.get_by_shortcut(sc_all[3][0])
    assert cmd
    # most all key cmd cannot be retrieved with getKeyEventsByCommand()
    # see: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1ui_1_1XAcceleratorConfiguration.html
    # This is an issue with LibreOffice. Not sure if it is a bug or a feature.
    # ShortCuts implements a work around for this.
    short_cuts = sc.get_by_command(cmd)
    assert short_cuts

    cmd = sc.get_by_shortcut(sc_all[4][0])
    assert cmd
    short_cuts = sc.get_by_command(cmd)
    assert short_cuts

    copy_short_cuts = sc.get_by_command(".uno:Copy")
    assert copy_short_cuts


def test_shortcut_calc(loader) -> None:
    # Test adding controls to a cell and a range
    # The controls are found when calling "cell.control.current_control" by shape position.
    from ooodev.calc import CalcDoc

    doc = CalcDoc.create_doc(loader)
    try:
        sc = Shortcuts(app=Service.CALC)
        sc_all = sc.get_all()
        assert sc_all
        cmd = sc.get_by_shortcut(sc_all[3][0])
        assert cmd
        short_cuts = sc.get_by_command(cmd)
        assert short_cuts
    finally:
        doc.close()
