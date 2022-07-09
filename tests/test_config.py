from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.cfg.config import Config

def test_config() -> None:
    c1 = Config()
    c2 = Config()
    assert c1 is c2
    assert Config().profile_versions[0] == "4"