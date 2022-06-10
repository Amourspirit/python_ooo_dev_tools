#!/usr/bin/env python
# coding: utf-8
import os
import sys
from subprocess import run


def main():
    myenv = os.environ.copy()
    pypath = ''
    p_sep = ';' if os.name == 'nt' else ':'
    for d in sys.path:
        pypath = pypath + d + p_sep
    myenv['PYTHONPATH'] = pypath
    cmd_str = 'setup.py clean --all'
    cmd = [sys.executable] + cmd_str.split()
    res = run(cmd, env=myenv)
    if res.stdout:
        print(res.stdout)
    if res.stderr:
        print(res.stderr)
    cmd_str = 'setup.py sdist bdist_wheel'
    cmd = [sys.executable] + cmd_str.split()
    res = run(cmd, env=myenv)
    if res.stdout:
        print(res.stdout)
    if res.stderr:
        print(res.stderr)
    cmd_str = '-m twine check dist/*'
    cmd = [sys.executable] + cmd_str.split()
    res = run(cmd, env=myenv)
    if res.stdout:
        print(res.stdout)
    if res.stderr:
        print(res.stderr)

if __name__ == '__main__':
    main()