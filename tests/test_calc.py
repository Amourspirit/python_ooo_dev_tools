import pytest
from pathlib import Path

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

def test_get_sheet() -> None:
    # get_sheet is overload method.
    # testiing each overload.
    from ooodev.utils import lo as mLo
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI
    Lo = mLo.Lo
    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None
        # GUI.set_visible(is_visible=True, odoc=doc)
        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 1
        assert sheet_names[0] == 'Sheet1'
        # test overloads
        sheet_1_1 = Calc.get_sheet(doc=doc,sheet_name='Sheet1')
        assert sheet_1_1 is not None
        name_1_1 = Calc.get_sheet_name(sheet_1_1)
        assert name_1_1 == 'Sheet1'
        
        sheet_1_2 = Calc.get_sheet(doc,'Sheet1')
        assert sheet_1_2 is not None
        name_1_2 = Calc.get_sheet_name(sheet_1_2)
        assert name_1_1 == name_1_2
        
        sheet_1_3 = Calc.get_sheet(doc=doc, index=0)
        assert sheet_1_3 is not None
        name_1_3 = Calc.get_sheet_name(sheet_1_3)
        assert name_1_3 == name_1_1
        
        sheet_1_4 = Calc.get_sheet(doc, 0)
        assert sheet_1_4 is not None
        name_1_4 = Calc.get_sheet_name(sheet_1_4)
        assert name_1_4 == name_1_1
        # Lo.delay(2000)
        Lo.close_doc(doc=doc, deliver_ownership=False)

def test_insert_sheet() -> None:
    from ooodev.utils import lo as mLo
    from ooodev.office.calc import Calc
    Lo = mLo.Lo
    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None
        
        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 1
        assert sheet_names[0] == 'Sheet1'
        
        new_sht = Calc.insert_sheet(doc, "mysheet", 1)
        assert new_sht is not None
        
        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 2
        assert sheet_names[1] == 'mysheet'

def test_remove_sheet() -> None:
    from ooodev.utils import lo as mLo
    from ooodev.office.calc import Calc
    Lo = mLo.Lo
    sheet_name = 'mysheet'
    def add_sheet(doc):
        new_sht = Calc.insert_sheet(doc, sheet_name, 1)
        return new_sht
    with Lo.Loader() as loader:
        assert loader is not None
        doc = Calc.create_doc(loader)
        assert doc is not None
        
        sheet_names = Calc.get_sheet_names(doc)
        
        new_sht = add_sheet(doc)
        assert new_sht is not None
    
        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 2
        assert sheet_names[1] == sheet_name
        # test overloads
        # test named args
        assert Calc.remove_sheet(doc=doc, sheet_name=sheet_name)
        
        new_sht = add_sheet(doc)
        assert new_sht is not None
        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 2
        assert sheet_names[1] == sheet_name
        # test positional args
        assert Calc.remove_sheet(doc, sheet_name)
        
        new_sht = add_sheet(doc)
        assert new_sht is not None
        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 2
        assert sheet_names[1] == sheet_name
        # test named args
        assert Calc.remove_sheet(doc=doc, index=1)
        
        new_sht = add_sheet(doc)
        assert new_sht is not None
        sheet_names = Calc.get_sheet_names(doc)
        assert len(sheet_names) == 2
        assert sheet_names[1] == sheet_name
        # test positional args
        assert Calc.remove_sheet(doc, 1)
        
        # sheet no longer exist
        assert Calc.remove_sheet(doc, 1) is False
        assert Calc.remove_sheet(doc, sheet_name) is False
        
        # incorrect number of arguments
        assert Calc.remove_sheet(doc) is False
        assert Calc.remove_sheet(doc, 0, "any") is False
        
        # incorrect type
        assert Calc.remove_sheet(doc, 22.3) is False
        
        