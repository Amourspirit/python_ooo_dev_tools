#!/usr/bin/env python
# coding: utf-8
import os
import sys
import subprocess
from typing import Any
from pathlib import Path
from ..utils import util
# from ..lib.connect import LoSocketStart
from ooodev.utils import uno_util


def run_lo_py(*args: str) -> None:
    """
    Runs a command.
    
    In Windows the command is run using LibreOffice ``python.exe``.
    
    In Linux the current active python is used.
    In this context it is usually the python from the virtual environment setup for this project.

    Arguments:
        args (str): args that are to be passed to the file being executed.
            The first arg must be the path to the pyton file being exectued
            such as 'C:\\Users\\user\\Documents\\Projects\\Python\\examplex\\src\\test.py'.

    Raises:
        FileNotFoundError: If file to run is not found.
        ValueError: If no arg is passed in.
        ValueError: If file arg is not an actual file.
    """
    if len(args) == 0:
        raise ValueError('No args to run. Must be at least one arg.')
    pargs = list(args)
    pfile = Path(pargs.pop(0))
    if not pfile.is_absolute():
        root = Path(os.environ["project_root"])
        pfile = Path(root, pfile)
    if not pfile.exists():
        raise FileNotFoundError(f"Unable to find '{pfile}'")
    if not pfile.is_file():
        raise ValueError(f"Not a file: '{pfile}'")
    cmd = [f'"{util.get_lo_python_ex()}"', f"{pfile}"] + pargs
    myenv = os.environ.copy()
    # print("cmd:", cmd)
    pypath = ''
    if sys.platform == 'win32':
        p_inst = uno_util.get_soffice_install_path()
        pypath = pypath + str(util.get_site_packeges_dir()) + ';'
        pypath = pypath + util.get_root() + ';'
        myenv['URE_BOOTSTRAP'] = f"vnd.sun.star.pathname:{p_inst}\\program\\fundamental.ini"
        myenv['UNO_PATH'] = f"{p_inst}\\program\\"
    myenv['PYTHONPATH'] = pypath
    # subprocess.run(cmd, env=myenv, ) fails in windows with error: PermissionError: [WinError 5] Access is denied
    # this is rather strange becuse it runs fine in debug mode.
    # solution is to run in shell.
    cmd_str = " ".join(cmd)
    subprocess.run(cmd_str, env=myenv, shell=True)
    


def run_py(*args: str) -> None:
    """
    Runs a command using the current active python.
  
    In this context it is usually the python from the virtual environment setup for this project.

    Arguments:
        args (str): args that are to be passed to the file being executed.
            The first arg must be the path to the pyton file being exectued
            such as 'C:\\Users\\user\\Documents\\Projects\\Python\\examplex\\src\\test.py'.

    Raises:
        FileNotFoundError: If file to run is not found.
        ValueError: If no arg is passed in.
        ValueError: If file arg is not an actual file.
    """
    if len(args) == 0:
        raise ValueError('No args to run. Must be at least one arg.')
    pargs = list(args)
    pfile = Path(pargs.pop(0))
    if not pfile.is_absolute():
        root = Path(os.environ["project_root"])
        pfile = Path(root, pfile)
    if not pfile.exists():
        raise FileNotFoundError(f"Unable to find '{pfile}'")
    if not pfile.is_file():
        raise ValueError(f"Not a file: '{pfile}'")
    myenv = os.environ.copy()
    pypath = ''
    p_sep = ';' if sys.platform == 'win32' else ':'
    for d in sys.path:
        pypath = pypath + d + p_sep
    if sys.platform == 'win32':
        p_inst = uno_util.get_soffice_install_path()
        pypath = str(Path(p_inst, 'program')) + ';' + pypath
        myenv['URE_BOOTSTRAP'] = f"vnd.sun.star.pathname:{p_inst}\\program\\fundamental.ini"
        myenv['UNO_PATH'] = f"{p_inst}\\program\\"
    # else:
        # pypath = util.get_root() + ';' + pypath
    myenv['PYTHONPATH'] = pypath
    cmd = [sys.executable, f"{pfile}"] + pargs
    # print("cmd:", cmd)
    process = subprocess.run(" ".join(cmd),  env=myenv, shell=True)
    # for c in iter(lambda: process.stdout.read(1), b''): 
    #     sys.stdout.buffer.write(c)