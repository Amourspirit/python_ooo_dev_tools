#!coding: utf-8
# python -msphinx.ext.intersphinx
from subprocess import run
import pathlib
import os
import sys

ROOT_PATH = pathlib.Path(__file__).parent.parent.parent
DOCS_PATH = ROOT_PATH / 'docs'


def main():
    global ROOT_PATH
    global DOCS_PATH
    myenv = os.environ.copy()
    pypath = ''
    p_sep = ';' if os.name == 'nt' else ':'
    for d in sys.path:
        pypath = pypath + d + p_sep
    myenv['PYTHONPATH'] = pypath

    obj_inv_path = DOCS_PATH / "_build" / "html" / "objects.inv"

    cmd_str = f"{sys.executable} -msphinx.ext.intersphinx {obj_inv_path}"
    res = run(cmd_str.split(), env=myenv)
    if res and res.returncode != 0:
        print(res)


if __name__ == '__main__':
    main()
