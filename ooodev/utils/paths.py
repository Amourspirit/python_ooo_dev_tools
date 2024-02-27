# coding: utf-8
from __future__ import annotations
from typing import Union
import os
import sys
import shutil
import __main__
from pathlib import Path
from typing import overload
from ooodev.utils.sys_info import SysInfo

# do not import from type_var here.
# this module is used by uno_lnk.py and type_var imports uno
# PathOrStr = Union[str, os.PathLike]

# python path on mac:  /Applications/LibreOffice.app/Contents/Resources/python
#   https://ask.libreoffice.org/t/where-is-the-python-executable-embedded-in-libreoffice-on-macos/50042


PLATFORM = SysInfo.get_platform()
if PLATFORM == SysInfo.PlatformEnum.WINDOWS:
    import winreg

_INSTALL_PATH = None


def get_soffice_install_path() -> Path:
    """
    Gets the Soffice install path.

    For windows this will be something like: ``C:\\Program Files\\LibreOffice``.
    For Linux this will be something like: ``/usr/lib/libreoffice``

    Returns:
        Path: install as Path.
    """
    # sourcery skip: low-code-quality
    global _INSTALL_PATH
    if _INSTALL_PATH is not None:
        return _INSTALL_PATH
    if PLATFORM == SysInfo.PlatformEnum.WINDOWS:
        # get the path location from registry
        value = ""
        for _key in (
            # LibreOffice 3.4.5,6,7 on Windows
            "SOFTWARE\\LibreOffice\\UNO\\InstallPath",
            # OpenOffice 3.3
            "SOFTWARE\\OpenOffice.org\\UNO\\InstallPath",
        ):
            try:
                value = winreg.QueryValue(winreg.HKEY_LOCAL_MACHINE, _key)  # type: ignore
            except Exception as detail:
                value = ""
                _errMess = f"{detail}"
            else:
                break  # first existing key will do
        if value != "":
            _INSTALL_PATH = Path("\\".join(value.split("\\")[:-1]))  # drop the program
            return _INSTALL_PATH

        # failed to get path from registry. Going Manual
        soffice = "soffice.exe"
        p_sf = Path(os.environ["PROGRAMFILES"], "LibreOffice", "program", soffice)
        if p_sf.exists() is False or p_sf.is_file() is False:
            p_sf = Path(os.environ["PROGRAMFILES(X86)"], "LibreOffice", "program", soffice)
        if p_sf.exists() is False or p_sf.is_file() is False:
            # perhaps running a developer version.
            # C:\Program Files\LibreOfficeDev 7\program
            p_sf = Path(os.environ["PROGRAMFILES"], "LibreOfficeDev 7", "program", soffice)
        if p_sf.exists() is False or p_sf.is_file() is False:
            p_sf = Path(os.environ["PROGRAMFILES(X86)"], "LibreOfficeDev 7", "program", soffice)
        if not p_sf.exists():
            raise FileNotFoundError(f"LibreOffice '{p_sf}' not found.")
        if not p_sf.is_file():
            raise IsADirectoryError(f"LibreOffice '{p_sf}' is not a file.")
        # drop \program\soffice.exe
        # expect C:\Program Files\LibreOffice
        _INSTALL_PATH = p_sf.parent.parent
        return _INSTALL_PATH

    elif PLATFORM == SysInfo.PlatformEnum.MAC:
        _INSTALL_PATH = Path("/Applications/LibreOffice.app/Contents/MacOS")
        return _INSTALL_PATH
    else:
        # unix
        soffice = "soffice"
        # search system path
        s = shutil.which(soffice)
        p_sf = None
        if s is not None:
            # expect '/usr/bin/soffice'
            if os.path.islink(s):
                p_sf = Path(os.path.realpath(s)).parent
            else:
                p_sf = Path(s).parent
        if p_sf is None:
            s = "/usr/bin/soffice"
            if os.path.islink(s):
                p_sf = Path(os.path.realpath(s)).parent
            else:
                p_sf = Path(s).parent
        if not p_sf.exists():
            raise FileNotFoundError(f"LibreOffice '{p_sf}' not found.")
        if not p_sf.is_file():
            raise IsADirectoryError(f"LibreOffice '{p_sf}' is not a file.")
        # drop /program/soffice
        _INSTALL_PATH = p_sf.parent.parent
        return _INSTALL_PATH


def get_soffice_path(check_env: bool = True) -> Path:
    """
    Gets path to soffice

    Args:
        check_env (bool, optional): If True then check environment variable ``ODEV_CONN_SOFFICE`` for path. Defaults to True.

    Returns:
        Path: path to soffice

    .. versionchanged:: 0.11.6
        Added ``check_env`` parameter.
    """
    if check_env and "ODEV_CONN_SOFFICE" in os.environ:
        return Path(os.environ["ODEV_CONN_SOFFICE"])
    if PLATFORM == SysInfo.PlatformEnum.WINDOWS:
        return Path(get_lo_path(), "soffice.exe")
    return Path(get_lo_path(), "soffice")


def get_uno_path() -> Path:
    """
    Searches known paths for path that contains uno.py

    This path is different for windows and linux.
    Typically for Windows ``C:\\Program Files\\LibreOffice\\program``
    Typically for Linux ``/usr/lib/python3/dist-packages``

    Raises:
        FileNotFoundError: if path is not found
        NotADirectoryError: if path is not a directory

    Returns:
        Path: First found path.
    """
    if PLATFORM == SysInfo.PlatformEnum.WINDOWS:
        p_uno = Path(os.environ["PROGRAMFILES"], "LibreOffice", "program")
        if not p_uno.exists() or not p_uno.is_dir():
            p_uno = Path(os.environ["PROGRAMFILES(X86)"], "LibreOffice", "program")
        if not p_uno.exists():
            raise FileNotFoundError("Uno Source Dir not found.")
        if not p_uno.is_dir():
            raise NotADirectoryError("UNO source is not a Directory")
        return p_uno
    elif PLATFORM == SysInfo.PlatformEnum.MAC:
        return Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    else:
        p_uno = Path("/usr/lib/python3/dist-packages")
        if not p_uno.exists():
            raise FileNotFoundError("Uno Source Dir not found.")
        if not p_uno.is_dir():
            raise NotADirectoryError("UNO source is not a Directory")
        return p_uno


def get_lo_path() -> Path:
    """
    Searches known paths for path that contains LibreOffice ``soffice``.

    This path is different for windows and linux.
    Typically for Windows ``C:\\Program Files\\LibreOffice\\program``
    Typically for Linux ``/usr/bin/soffice``

    Raises:
        FileNotFoundError: if path is not found
        NotADirectoryError: if path is not a directory

    Returns:
        Path: First found path.
    """
    if PLATFORM == SysInfo.PlatformEnum.WINDOWS:
        return Path(get_soffice_install_path(), "program")

    elif PLATFORM == SysInfo.PlatformEnum.MAC:
        return Path("/Applications/LibreOffice.app/Contents/MacOS")
    else:
        # search system path
        s = shutil.which("soffice")
        p_sf = None
        if s is not None:
            # expect '/usr/bin/soffice'
            if os.path.islink(s):
                # follow link
                p_sf = Path(os.path.realpath(s)).parent
            else:
                p_sf = Path(s).parent
        if p_sf is None:
            p_sf = Path("/usr/bin/soffice")
            if not p_sf.exists() or not p_sf.is_file():
                raise FileNotFoundError("LibreOffice Source Dir not found.")
            p_sf = p_sf.parent

        if not p_sf.exists():
            raise FileNotFoundError("LibreOffice Source Dir not found.")
        if not p_sf.is_dir():
            raise NotADirectoryError("LibreOffice source is not a Directory")
        return p_sf


def get_lo_python_ex() -> str:
    """
    Gets the python executable for different environments.

    In Linux this is the current python executable.
    If a virtual environment is activated then that will be the
    python executable that is returned.

    In Windows this is the ``python.exe`` file in LibreOffice.
    Typically for Windows ``C:\\Program Files\\LibreOffice\\program\\python.exe``

    Raises:
        FileNotFoundError: In Windows if ``python.exe`` is not found.
        NotADirectoryError: In Windows if ``python.exe`` is not a file.

    Returns:
        str: file location of python executable.
    """
    if PLATFORM != SysInfo.PlatformEnum.WINDOWS:
        return sys.executable
    p = Path(get_lo_path(), "python.exe")

    if not p.exists():
        raise FileNotFoundError("LibreOffice python executable not found.")
    if not p.is_file():
        raise NotADirectoryError("LibreOffice  python executable is not a file")
    return str(p)


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


def get_virtual_env_site_packages_path() -> str:
    """
    Gets the Virtual Environment ``site_packages`` Path

    Returns:
        str: Virtual Environment Path

    Note:
        If unable to get virtual path from Environment then ``sys.base_exec_prefix`` is returned.

    .. versionadded:: 0.11.7
    """
    s_path = os.environ.get("VIRTUAL_ENV", None)
    sys_ver = f"python{sys.version_info[0]}.{sys.version_info[1]}"
    if s_path is not None:
        return os.path.join(s_path, "lib", sys_ver, "site-packages")
    else:
        return os.path.join(sys.base_exec_prefix, "lib", sys_ver, "site-packages")


@overload
def mkdirp(dest_dir: str) -> None: ...


@overload
def mkdirp(dest_dir: Path) -> None: ...


def mkdirp(dest_dir: Union[str, os.PathLike]) -> None:
    """
    Creates path and sub paths not existing.

    Args:
        dest_dir (str | PathLike]): PathLike object
    """
    # Python ≥ 3.5
    pth = Path(dest_dir)
    if not pth.is_absolute():
        pth = pth.resolve()
    pth.mkdir(parents=True, exist_ok=True)
