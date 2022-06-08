# coding: utf-8
import os
import sys
import shutil
import __main__
from pathlib import Path
from typing import Union, overload

if sys.platform == "win32":
    import winreg

_INSTALL_PATH = None


def get_soffice_install_path() -> Path:
    """
    Gets the Soffice install path.

    For windows this will be something like: ``C:\Program Files\LibreOffice``.
    For Linux this will be somethon like: ``/usr/lib/libreoffice``

    Returns:
        Path: install as Path.
    """
    global _INSTALL_PATH
    if _INSTALL_PATH is not None:
        return _INSTALL_PATH
    if sys.platform == "win32":
        # get the path location from registery
        value = ""
        for _key in (
            # LibreOffice 3.4.5,6,7 on Windows
            "SOFTWARE\\LibreOffice\\UNO\\InstallPath",
            # OpenOffice 3.3
            "SOFTWARE\\OpenOffice.org\\UNO\\InstallPath",
        ):
            try:
                value = winreg.QueryValue(winreg.HKEY_LOCAL_MACHINE, _key)
            except Exception as detail:
                value = ""
                _errMess = "%s" % detail
            else:
                break  # first existing key will do
        if value != "":
            _INSTALL_PATH = Path("\\".join(value.split("\\")[:-1]))  # drop the program
            return _INSTALL_PATH

        # failed to get path from registery. Going Manual
        soffice = "soffice.exe"
        p_sf = Path(os.environ["PROGRAMFILES"], "LibreOffice", "program", soffice)
        if p_sf.exists() is False or p_sf.is_file() is False:
            p_sf = Path(os.environ["PROGRAMFILES(X86)"], "LibreOffice", "program", soffice)
        if not p_sf.exists():
            raise FileNotFoundError(f"LibreOffice '{p_sf}' not found.")
        if not p_sf.is_file():
            raise IsADirectoryError(f"LibreOffice '{p_sf}' is not a file.")
        # drop \program\soffice.exe
        # expect C:\Program Files\LibreOffice
        _INSTALL_PATH = p_sf.parent.parent
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
                p_sf = Path(os.path.realpath(s))
            else:
                p_sf = Path(s)
        if p_sf is None:
            s = "/usr/bin/soffice"
            if os.path.islink(s):
                p_sf = Path(os.path.realpath(s))
            else:
                p_sf = Path(s)
        if not p_sf.exists():
            raise FileNotFoundError(f"LibreOffice '{p_sf}' not found.")
        if not p_sf.is_file():
            raise IsADirectoryError(f"LibreOffice '{p_sf}' is not a file.")
        # drop /program/soffice
        _INSTALL_PATH = p_sf.parent.parent
        return _INSTALL_PATH


def get_uno_path() -> Path:
    """
    Searches known paths for path that contains uno.py

    This path is different for windows and linux.
    Typically for Windows ``C:\Program Files\LibreOffice\program``
    Typically for Linux ``/usr/lib/python3/dist-packages``

    Raises:
        FileNotFoundError: if path is not found
        NotADirectoryError: if path is not a directory

    Returns:
        Path: First found path.
    """
    if os.name == "nt":

        p_uno = Path(os.environ["PROGRAMFILES"], "LibreOffice", "program")
        if p_uno.exists() is False or p_uno.is_dir() is False:
            p_uno = Path(os.environ["PROGRAMFILES(X86)"], "LibreOffice", "program")
        if not p_uno.exists():
            raise FileNotFoundError("Uno Source Dir not found.")
        if not p_uno.is_dir():
            raise NotADirectoryError("UNO source is not a Directory")
        return p_uno
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
    Typically for Windows ``C:\Program Files\LibreOffice\program``
    Typically for Linux `/usr/bin/soffice``

    Raises:
        FileNotFoundError: if path is not found
        NotADirectoryError: if path is not a directory

    Returns:
        Path: First found path.
    """
    if os.name == "nt":

        p_uno = Path(os.environ["PROGRAMFILES"], "LibreOffice", "program")
        if p_uno.exists() is False or p_uno.is_dir() is False:
            p_uno = Path(os.environ["PROGRAMFILES(X86)"], "LibreOffice", "program")
        if not p_uno.exists():
            raise FileNotFoundError("LibreOffice Source Dir not found.")
        if not p_uno.is_dir():
            raise NotADirectoryError("LibreOffice source is not a Directory")
        return p_uno
    else:
        # search system path
        s = shutil.which("soffice")
        p_sf = None
        if s is not None:
            # expect '/usr/bin/soffice'
            if os.path.islink(s):
                p_sf = Path(os.path.realpath(s)).parent
            else:
                p_sf = Path(s).parent
        if p_sf is None:
            p_sf = Path("/usr/bin/soffice")
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
    python exceutable that is returned.

    In Windows this is the python.exe file in LibreOffice.
    Typically for Windows ``C:\Program Files\LibreOffice\program\python.exe``

    Raises:
        FileNotFoundError: In Windows if python.exe is not found.
        NotADirectoryError: In Windows if python.exe is not a file.

    Returns:
        str: file location of python executable.
    """
    if os.name == "nt":
        p = Path(get_lo_path(), "python.exe")

        if not p.exists():
            raise FileNotFoundError("LibreOffice python executable not found.")
        if not p.is_file():
            raise NotADirectoryError("LibreOffice  python executable is not a file")
        return str(p)
    else:
        return sys.executable


@overload
def mkdirp(dest_dir: str) -> None:
    ...


@overload
def mkdirp(dest_dir: Path) -> None:
    ...


def mkdirp(dest_dir: Union[str, Path]) -> None:
    """
    Creates path and subpaths not existing.

    Args:
        dest_dir (str | Path): PathLike object
    """
    # Python ≥ 3.5
    if isinstance(dest_dir, Path):
        dest_dir.mkdir(parents=True, exist_ok=True)
    else:
        Path(dest_dir).mkdir(parents=True, exist_ok=True)
