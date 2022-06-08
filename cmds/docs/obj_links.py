#!coding: utf-8
# python -msphinx.ext.intersphinx
from subprocess import run
import pathlib

ROOT_PATH = pathlib.Path(__file__).parent.parent.parent
DOCS_PATH = ROOT_PATH / 'docs'


def main():
    global ROOT_PATH
    global DOCS_PATH

    obj_inv_path = DOCS_PATH / "_build" / "html" / "objects.inv"

    cmd_str = f"python -msphinx.ext.intersphinx {obj_inv_path}"
    res = run(cmd_str.split())
    if res and res.returncode != 0:
        print(res)


if __name__ == '__main__':
    main()
