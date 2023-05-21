#!/usr/bin/env python
# coding: utf-8
"""
Creates system links for uno and uno_helper.

This module is not useful in Windows.

Windows makes use of the run_auto.py and creates subprocesses that call
LibreOffice's python.exe while setting the appropriate paths so LibreOffice python can access
this code base and its site-packages.
"""
import os
import sys
import shutil
from typing import Optional
from pathlib import Path
from ..utils import util
from ooodev.utils import paths


def add_links(uno_src_dir: Optional[str] = None):
    if isinstance(uno_src_dir, str):
        str_cln = uno_src_dir.strip()
        if len(str_cln) == 0:
            p_uno_dir = paths.get_uno_path()
        else:
            p_uno_dir = Path(str_cln)
            if not p_uno_dir.exists():
                raise FileNotFoundError(f"Uno Source Dir not found: {uno_src_dir}")
            if not p_uno_dir.is_dir():
                raise NotADirectoryError(f"UNO source is not a Directory: {uno_src_dir}")
    else:
        p_uno_dir = paths.get_uno_path()
    p_site_dir = util.get_site_packages_dir()
    if p_site_dir is None:
        print("Unable to find site_packages direct in virtual enviornment")
        return

    p_uno = Path(p_uno_dir, "uno.py")
    p_uno_helper = Path(p_uno_dir, "unohelper.py")

    if p_uno.exists():
        dest = Path(p_site_dir, "uno.py")
        try:
            os.symlink(src=p_uno, dst=dest)
            print(f"Created system link: {p_uno} -> {dest}")
        except FileExistsError:
            print(f"File already exist: {dest}")
        except OSError:
            # OSError: [WinError 1314] A required privilege is not held by the client
            print(f"Unable to create system link for  '{p_uno.name}'. Attempting copy.")
            shutil.copy2(p_uno, dest)
            print(f"Copied file: {p_uno} -> {dest}")
    else:
        print(f"{p_uno.name} not found.")

    if p_uno_helper.exists():
        dest = Path(p_site_dir, "unohelper.py")
        try:
            os.symlink(src=p_uno_helper, dst=dest)
            print(f"Created system link: {p_uno_helper} -> {dest}")
        except FileExistsError:
            print(f"File already exist: {dest}")
        except OSError:
            # OSError: [WinError 1314] A required privilege is not held by the client
            print(f"Unable to create system link for  '{p_uno_helper.name}'. Attempting copy.")
            shutil.copy2(p_uno_helper, dest)
            print(f"Copied file: {p_uno_helper} -> {dest}")
    else:
        print(f"{p_uno_helper.name} not found.")
    return
    p_scriptforge = Path(paths.get_lo_path(), "scriptforge.py")
    if p_scriptforge.exists():
        dest = Path(p_site_dir, "scriptforge.py")
        try:
            os.symlink(src=p_scriptforge, dst=dest)
            print(f"Created system link: {p_scriptforge} -> {dest}")
        except FileExistsError:
            print(f"File already exist: {dest}")
        except OSError:
            # OSError: [WinError 1314] A required privilege is not held by the client
            print(f"Unable to create system link for  '{p_scriptforge.name}'. Attempting copy.")
            shutil.copy2(p_scriptforge, dest)
            print(f"Copied file: {p_scriptforge} -> {dest}")
    else:
        print(f"{p_scriptforge.name} not found.")


def remove_links():
    p_site_dir = util.get_site_packages_dir()
    if p_site_dir is None:
        print("Unable to find site_packages direct in virtual enviornment")
        return

    uno_path = Path(p_site_dir, "uno.py")
    if uno_path.exists():
        os.remove(uno_path)
        print("removed uno.py")
    else:
        print("uno.py does not exist in virtual env.")
    unohelper_path = Path(p_site_dir, "unohelper.py")
    if unohelper_path.exists():
        os.remove(unohelper_path)
        print("removed unohelper.py")
    else:
        print("unohelper.py does not exist in virtual env.")
    return
    scriptforge_path = Path(p_site_dir, "scriptforge.py")
    if scriptforge_path.exists():
        os.remove(scriptforge_path)
        print("removed scriptforge.py")
    else:
        print("scriptforge.py does not exist in virtual env.")


def main():
    if len(sys.argv) == 2:
        arg = sys.argv[1]
        if arg == "-r" or arg == "--remove":
            remove_links()
            return
        if arg == "-a" or arg == "-add":
            add_links()
            return
    print("for add links use -a or --add\nfor remove use -r or --remove")


if __name__ == "__main__":
    main()
