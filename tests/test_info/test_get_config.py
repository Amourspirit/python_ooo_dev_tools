import pytest

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])



def test_config(loader) -> None:
    from ooodev.utils.info import Info
    name = Info.get_config('ooName')
    assert name == 'LibreOffice'
    
    addin = Info.get_paths('Addin')
    assert addin is not None
    
    op = Info.get_office_dir()
    assert op is not None