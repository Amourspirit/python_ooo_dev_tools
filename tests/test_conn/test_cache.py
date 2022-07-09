from unittest.mock import patch
from pathlib import Path
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.conn.cache import Cache
from ooodev.utils.sys_info import SysInfo

def test_cache_path_linux_mac() -> None:
    with patch.object(Path, 'exists') as mock_exists:
        mock_exists.return_value = True
        with patch.object(Path, 'is_dir') as mock_is_dir:
            mock_is_dir.return_value = True
            with patch.object(SysInfo, 'get_platform') as mock_platform:
                mock_platform.return_value = SysInfo.PlatformEnum.MAC
                c = Cache()
                assert c.cache_path is not None
            
            with patch.object(SysInfo, 'get_platform') as mock_platform:
                mock_platform.return_value = SysInfo.PlatformEnum.LINUX
                c = Cache()
                assert c.cache_path is not None

def test_cache_path_win(monkeypatch) -> None:
    monkeypatch.setenv('APPDATA', 'C:\\Users\\test\\AppData\\Roaming')
    with patch.object(Path, 'exists') as mock_exists:
        mock_exists.return_value = True
        with patch.object(Path, 'is_dir') as mock_is_dir:
            mock_is_dir.return_value = True
         
            with patch.object(SysInfo, 'get_platform') as mock_platform:
                mock_platform.return_value = SysInfo.PlatformEnum.WINDOWS
                c = Cache()
                assert c.cache_path is not None

def test_cache_path_no_exist() -> None:
    # p = Path("~/.config/libreoffice/4")
    with patch.object(Path, 'exists') as mock_exists:
        mock_exists.return_value = False
        # p.exists()
        c = Cache()
        assert c.cache_path is None

def test_copy_profile(tmp_path_fn:Path, fixture_path:Path) -> None:
    cache_path = Path(fixture_path, "dummy_profile")
    working_dir = Path(tmp_path_fn, "LO1")
    working_dir.mkdir()
    c = Cache(working_dir=working_dir, cache_path=cache_path)
    c.copy_cache_to_profile()
    tmp_path = Path(working_dir, "profile/4/user/data.txt")
    assert tmp_path.exists()
    assert tmp_path.is_file()
    assert c._profile_cached is True
    
    cache_path2 = Path(tmp_path_fn, "LO2")
    c2 = Cache(cache_path=cache_path2, working_dir=working_dir)
    c2.cache_profile()
    tmp_path2 = Path(cache_path2, "4/user/data.txt")
    assert tmp_path2.exists()
    assert tmp_path2.is_file()
    
    c.del_working_dir()
    assert working_dir.exists() == False