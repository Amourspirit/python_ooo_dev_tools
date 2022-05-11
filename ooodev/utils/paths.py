# coding: utf-8
import os
import sys
import shutil
import __main__
from pathlib import Path
from typing import Union, overload


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
        dest_dir (Union[str, Path]): PathLike object

    Since:
        Python 3.5
    """
    # Python ≥ 3.5
    if isinstance(dest_dir, Path):
        dest_dir.mkdir(parents=True, exist_ok=True)
    else:
        Path(dest_dir).mkdir(parents=True, exist_ok=True)
