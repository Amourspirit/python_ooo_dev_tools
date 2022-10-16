from __future__ import annotations
from typing import Tuple
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.cfg.config import Config
from pytest import MonkeyPatch

def test_config(monkeypatch: MonkeyPatch) -> None:
    monkeypatch.setattr(Config, "_instance", None)
    c1 = Config()
    c2 = Config()
    assert c1 is c2
    assert Config().profile_versions[0] == "4"


@pytest.mark.parametrize("expected,env_val", [(("3",), "3"), (("4", "5"), "4,5"), (("3", "4", "5"), "3,4, 5")])
def test_config_env_profile_version(expected: Tuple[str, ...], env_val: str, monkeypatch: MonkeyPatch) -> None:
    monkeypatch.setattr(Config, "_instance", None)
    monkeypatch.setenv(name="OOODEV_CONFIG_PROFILE_VERSION", value=env_val)
    exp = set(expected)
    acutual = set(Config().profile_versions)
    assert exp == acutual

@pytest.mark.parametrize("env_val", [("/mypath/some/path"), ("\\some\\path\\other")])
def test_config_env_slide_template_path(env_val: str, monkeypatch: MonkeyPatch) -> None:
    monkeypatch.setattr(Config, "_instance", None)
    monkeypatch.setenv(name="OOODEV_CONFIG_SLIDE_TEMPLATE_PATH", value=env_val)
    assert Config().slide_template_path == env_val