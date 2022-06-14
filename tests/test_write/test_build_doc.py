import os
import pytest
from pathlib import Path
# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

def test_build_doc(loader,props_str_to_dict, capsys: pytest.CaptureFixture):
    from ooodev.utils.lo import Lo
    from ooodev.office.write import Write
    from ooodev.utils.gui import GUI
    from ooodev.utils.props import Props
    visible = True
    delay = 1000
    doc = Write.create_doc(loader)
    if visible:
        GUI.set_visible(visible, doc)
    
    cursor = Write.get_cursor(doc)
    capsys.readouterr() # clear buffer
    Props.show_obj_props(prop_kind="Cursor", obj=cursor)
    cap = capsys.readouterr()
    cap_out = cap.out
    assert cap_out is not None
    prop_dict: dict = props_str_to_dict(cap_out)
    assert prop_dict["ParaBackTransparent"] == "True"
    assert prop_dict["Endnote"] == "None"
    assert prop_dict["Footnote"] == "None"
    assert prop_dict["TextField"] == "None"
    assert prop_dict["TextTable"] == "None"
    
    Write.append(cursor=cursor, text="Some examples of simple text ")
    Write.append(cursor, "styles.")
    Write.append(cursor=cursor,ctl_char=Write.ControlCharacter.LINE_BREAK)
    
    Lo.delay(delay)
    Lo.close_doc(doc, False)