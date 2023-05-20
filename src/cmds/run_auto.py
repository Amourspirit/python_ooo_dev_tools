from __future__ import annotations
import os
import sys
import subprocess
from pathlib import Path
from ..utils import util

from oooenv.utils import uno_paths
from oooenv.utils import local_paths


def run_lo_py(*args: str) -> None:
    """
    Runs a command.

    In Windows the command is run using LibreOffice ``python.exe``.

    In Linux the current active python is used.
    In this context it is usually the python from the virtual environment setup for this project.

    Arguments:
        args (str): args that are to be passed to the file being executed.
            The first arg must be the path to the python file being executed
            such as 'C:\\Users\\user\\Documents\\Projects\\Python\\examples\\src\\test.py'.

    Raises:
        FileNotFoundError: If file to run is not found.
        ValueError: If no arg is passed in.
        ValueError: If file arg is not an actual file.
    """
    if not args:
        raise ValueError("No args to run. Must be at least one arg.")
    p_args = list(args)
    p_file = Path(p_args.pop(0))
    if not p_file.is_absolute():
        root = Path(os.environ["project_root"])
        p_file = Path(root, p_file)
    if not p_file.exists():
        raise FileNotFoundError(f"Unable to find '{p_file}'")
    if not p_file.is_file():
        raise ValueError(f"Not a file: '{p_file}'")
    cmd = [f'"{uno_paths.get_lo_python_ex()}"', f"{p_file}"] + p_args
    my_env = os.environ.copy()
    # print("cmd:", cmd)
    py_path = ""
    if sys.platform == "win32":
        p_inst = uno_paths.get_soffice_install_path()
        py_path = py_path + str(local_paths.get_site_packeges_dir()) + ";"
        py_path = py_path + util.get_root() + ";"
        my_env["URE_BOOTSTRAP"] = f"vnd.sun.star.pathname:{p_inst}\\program\\fundamental.ini"
        my_env["UNO_PATH"] = f"{p_inst}\\program\\"
    my_env["PYTHONPATH"] = py_path
    # subprocess.run(cmd, env=my_env, ) fails in windows with error: PermissionError: [WinError 5] Access is denied
    # this is rather strange becuase it runs fine in debug mode.
    # solution is to run in shell.
    cmd_str = " ".join(cmd)
    subprocess.run(cmd_str, env=my_env, shell=True)


def run_py(*args: str) -> None:
    """
    Runs a command using the current active python.

    In this context it is usually the python from the virtual environment setup for this project.

    Arguments:
        args (str): args that are to be passed to the file being executed.
            The first arg must be the path to the python file being executed
            such as 'C:\\Users\\user\\Documents\\Projects\\Python\\examples\\src\\test.py'.

    Raises:
        FileNotFoundError: If file to run is not found.
        ValueError: If no arg is passed in.
        ValueError: If file arg is not an actual file.
    """
    if not args:
        raise ValueError("No args to run. Must be at least one arg.")
    p_args = list(args)
    p_file = Path(p_args.pop(0))
    if not p_file.is_absolute():
        root = Path(os.environ["project_root"])
        p_file = Path(root, p_file)
    if not p_file.exists():
        raise FileNotFoundError(f"Unable to find '{p_file}'")
    if not p_file.is_file():
        raise ValueError(f"Not a file: '{p_file}'")
    my_env = os.environ.copy()
    py_path = ""
    p_sep = ";" if sys.platform == "win32" else ":"
    for d in sys.path:
        py_path = py_path + d + p_sep
    if sys.platform == "win32":
        p_inst = uno_paths.get_soffice_install_path()
        py_path = str(Path(p_inst, "program")) + ";" + py_path
        my_env["URE_BOOTSTRAP"] = f"vnd.sun.star.pathname:{p_inst}\\program\\fundamental.ini"
        my_env["UNO_PATH"] = f"{p_inst}\\program\\"
    # else:
    # py_path = util.get_root() + ';' + py_path
    my_env["PYTHONPATH"] = py_path
    cmd = [sys.executable, f"{p_file}"] + p_args
    # print("cmd:", cmd)
    process = subprocess.run(" ".join(cmd), env=my_env, shell=True)
    # for c in iter(lambda: process.stdout.read(1), b''):
    #     sys.stdout.buffer.write(c)
