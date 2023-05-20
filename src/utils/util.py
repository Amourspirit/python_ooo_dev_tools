from __future__ import annotations
import os
import sys
import shutil
import __main__
from pathlib import Path
from typing import List, Union, overload

_APP_ROOT = None
_OS_PATH_SET = False


def get_root() -> str:
    """
    Gets Application Root Path

    Returns:
        str: App root as string.
    """
    global _APP_ROOT
    if _APP_ROOT is None:
        _APP_ROOT = os.environ.get("project_root", str(Path(__main__.__file__).parent))
    return _APP_ROOT


def set_os_root_path() -> None:
    """
    Ensures application root dir is in sys.path
    """
    global _OS_PATH_SET
    if _OS_PATH_SET is False:
        _app_root = get_root()
        if _app_root not in sys.path:
            sys.path.insert(0, _app_root)
    _OS_PATH_SET = True


@overload
def get_path(path: str, ensure_absolute: bool = False) -> Path:
    ...


@overload
def get_path(path: List[str], ensure_absolute: bool = False) -> Path:
    ...


@overload
def get_path(path: Path, ensure_absolute: bool = False) -> Path:
    ...


def get_path(path: Union[str, Path, List[str]], ensure_absolute: bool = False) -> Path:
    """
    Builds a Path from a list of strings

    If path starts with ``~`` then it is expanded to user home dir.

    Args:
        lst (List[str], Path, str): List of path parts
        ensure_absolute (bool, optional): If true returned will have root dir prepended
            if path is not absolute

    Raises:
        ValueError: If lst is empty

    Returns:
        Path: Path of combined from ``lst``
    """
    p = None
    lst = []
    expand = None
    if isinstance(path, str):
        expand = path.startswith("~")
        p = Path(path)
    elif isinstance(path, Path):
        p = path
    else:
        lst = list(path)
    if p is None:
        if not lst:
            raise ValueError("lst arg is zero length")
        arg = lst[0]
        expand = arg.startswith("~")
        p = Path(*lst)
    elif expand is None:
        p_str = str(p)
        expand = p_str.startswith("~")
    if expand:
        p = p.expanduser()
    if ensure_absolute and p.is_absolute() is False:
        p = Path(get_root(), p)
    return p


def get_virtual_env_path() -> str:
    """
    Gets the Virtual Environment Path

    Returns:
        str: Virtual Environment Path

    Note:
        If unable to get virtual path from Environment then ``sys.base_exec_prefix`` is returned.
    """
    s_path = os.environ.get("VIRTUAL_ENV", None)
    return s_path if s_path is not None else sys.base_exec_prefix


def get_site_packages_dir() -> Union[Path, None]:
    """
    Gets the ``site-packages`` directory for current python environment.

    Returns:
        Union[Path, None]: site-packages dir if found; Otherwise, None.
    """
    v_path = get_virtual_env_path()
    p_site = Path(v_path, "Lib", "site-packages")
    if p_site.exists() and p_site.is_dir():
        return p_site

    ver = f"{sys.version_info[0]}.{sys.version_info[1]}"
    p_site = Path(v_path, "lib", f"python{ver}", "site-packages")
    return p_site if p_site.exists() and p_site.is_dir() else None


def copy_file(src: str | Path, dst: str | Path):
    shutil.copy2(src=src, dst=dst)
