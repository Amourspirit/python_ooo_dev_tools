# coding: utf-8
import os
import sys
from pathlib import Path
import shutil

if sys.platform == "win32":
    import winreg

_INSTALL_PATH = None


def get_soffice_install_path() -> Path:
    """
    Gets the Soffice install path.

    For windows this will be something like: ``C:\Program Files\LibreOffice``.
    For Linux this will be something like: ``/usr/lib/libreoffice``

    Returns:
        Path: install as Path.
    """
    global _INSTALL_PATH
    if _INSTALL_PATH is not None:
        return _INSTALL_PATH
    if sys.platform == "win32":
        # get the path location from registry
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

        # failed to get path from registry. Going Manual
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
