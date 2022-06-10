# coding: utf-8
import os
import sys
import subprocess


def main():
    myenv = os.environ.copy()
    pypath = ''
    p_sep = ';' if os.name == 'nt' else ':'
    for d in sys.path:
        pypath = pypath + d + p_sep
    myenv['PYTHONPATH'] = pypath
    cmd_str = 'twine upload --repository-url https://test.pypi.org/legacy/ dist/*'
    res = subprocess.run(cmd_str.split(), env=myenv)
    if res and res.returncode != 0:
        print(res)


if __name__ == '__main__':
    main()