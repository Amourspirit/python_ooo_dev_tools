#!/usr/bin/env python
import pathlib
from setuptools import setup, find_packages
from ooodev import __version__
PKG_NAME = 'ooo-dev-tools'
VERSION = __version__

# The directory containing this file
HERE = pathlib.Path(__file__).parent
# The text of the README file
with open(HERE / "README.rst") as fh:
    README = fh.read()

setup(
    name=PKG_NAME,
    version=VERSION,
    # package_data={"": ["*.json"]},
    python_requires='>=3.7.0',
    url="https://github.com/Amourspirit/python_ooo_dev_tools",
    packages=find_packages(exclude=['src', 'src.*', 'env', 'env.*', 'cmds', 'cmds.*']),
    author=":Barry-Thomas-Paul: Moss",
    author_email='bigbytetech@gmail.com',
    license="mit",
    keywords=['libreoffice', 'openoffice' 'macro', 'uno', 'ooouno', 'pyuno'],
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Environment :: Other Environment",
        "Intended Audience :: Developers",
        "Operating System :: OS Independent",
        "Topic :: Office/Business",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
    ],
    description="LibreOffice Developer Search Engine",
    long_description_content_type="text/rst",
    long_description=README
)