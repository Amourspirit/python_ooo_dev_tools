from __future__ import annotations
import os
import sys
import subprocess
from pathlib import Path
from typing import NamedTuple
from ..utils import util

# from ..lib.connect import LoSocketStart
from ooodev.utils import lo_util


class Version(NamedTuple):
    major: int
    minor: int
    revision: int

    def __str__(self) -> str:
        return f"{self.major}.{self.minor}.{self.revision}"


def get_uno_python_exe() -> str:
    """
    Gets Python Exe Path

    Raises:
        Exception: If not on Windows.
    """
    if sys.platform != "win32":
        raise Exception("Method only support Windows")
    p = Path(lo_util.get_soffice_install_path(), "program", "python.exe")
    return str(p)


def get_uno_python_ver() -> Version:
    """Gets Uno Python Version"""
    python_exe = get_uno_python_exe()
    output = subprocess.check_output([python_exe, "--version"]).decode("UTF8").strip()
    # somethink like Python 3.8.10
    parts = output.split()
    major, minor, rev = parts[1].split(".")
    # RangeParts()
    return Version(major=int(major), minor=int(minor), revision=int(rev))


def read_pyvenv_cfg(fnm: str = "pyvenv.cfg") -> dict:
    pyvenv_cfg = _get_pyvenv_cfg_path(fnm=fnm)
    result = {}
    with open(pyvenv_cfg, "r") as file:
        # strip of new line and remove anything after //
        # # for comment
        data = (row.partition("#")[0].rstrip() for row in file)
        # chain generator
        # remove empty lines
        data = (row for row in data if row)
        # each line should now be key valu pairs seperated by =
        for row in data:
            key, value = row.split("=")
            result[key.strip()] = value.strip()
    return result


def is_env_uno_python(cfg: dict | None = None) -> bool:
    if cfg is None:
        cfg = read_pyvenv_cfg()
    home = cfg.get("home", "")
    if not home:
        return False
    lo_path = _get_lo_path()
    return home.lower() == lo_path.lower()


def backup_cfg() -> None:
    src = _get_pyvenv_cfg_path()
    cfg = read_pyvenv_cfg()
    if is_env_uno_python(cfg):
        dst = src.parent / "pyvenv_uno.cfg"
    else:
        dst = src.parent / "pyvenv_orig.cfg"
    util.copy_file(src=src, dst=dst)


def _save_config(cfg: dict, fnm: str = "pyvenv.cfg"):
    lst = []
    for k, v in cfg.items():
        lst.append(f"{k} = {v}")
    if len(lst) > 0:
        lst.append("")
    f_out = _get_venv_path() / fnm
    with open(f_out, "w") as file:
        file.write("\n".join(lst))
    print("Saved cfg")


def toggle_cfg(suffix: str = "") -> None:
    env_path = _get_venv_path()
    if suffix:
        src = env_path / f"pyvenv_{suffix.strip()}.cfg"
        if not src.exists():
            print('File not found: "{src}"')
            print("No action taken")
            return
        dst = env_path / "pyvenv.cfg"
        util.copy_file(src=src, dst=dst)
        print(f"Set to {suffix.strip()} environment.")
        return

    cfg = read_pyvenv_cfg()
    if is_env_uno_python(cfg):
        src = env_path / "pyvenv_orig.cfg"
        dst = env_path / "pyvenv.cfg"
        util.copy_file(src=src, dst=dst)
        print("Set to Original Environment")
        return

    src = env_path / "pyvenv_orig.cfg"
    if not src.exists():
        _save_config(cfg=cfg, fnm="pyvenv_orig.cfg")

    uno_cfg = env_path / "pyvenv_uno.cfg"
    if not uno_cfg.exists():
        # need to create the file.
        cfg["home"] = _get_lo_path()
        cfg["version"] = str(get_uno_python_ver())
        cfg["include-system-site-packages"] = "false"
        _save_config(cfg=cfg, fnm="pyvenv_uno.cfg")

    src = env_path / "pyvenv_uno.cfg"
    dst = env_path / "pyvenv.cfg"
    util.copy_file(src=src, dst=dst)
    print("Set to UNO Environment")


def _get_lo_path() -> str:
    lo_path = os.environ.get("ODEV_CONN_SOFFICE", None)
    if lo_path:
        index = lo_path.rfind("program")
        if index > -1:
            lo_path = lo_path[: index + 7]
        else:
            lo_path = None
    if not lo_path:
        lo_path = str(lo_util.get_soffice_install_path() / "program")
    return lo_path


def _get_venv_path() -> Path:
    vpath = os.environ.get("VIRTUAL_ENV", None)
    if vpath is None:
        raise ValueError("Unable to get Virtual Environment Path")
    return Path(vpath)


def _get_pyvenv_cfg_path(fnm: str = "pyvenv.cfg") -> Path:
    vpath = _get_venv_path()
    pyvenv_cfg = Path(vpath, fnm)
    if not pyvenv_cfg.exists():
        raise FileNotFoundError(str(pyvenv_cfg))
    if not pyvenv_cfg.is_file():
        raise Exception(f'Not a file: "{pyvenv_cfg}"')
    return pyvenv_cfg
