#!/usr/bin/env python
# coding: utf-8
import sys
from subprocess import run


def main():
    cmd_str = 'setup.py clean --all'
    cmd = [sys.executable] + cmd_str.split()
    res = run(cmd)
    if res.stdout:
        print(res.stdout)
    if res.stderr:
        print(res.stderr)


if __name__ == '__main__':
    main()