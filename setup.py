#!/usr/bin/env python
import pathlib
from setuptools import setup, find_packages
from ooodev import __version__

PKG_NAME = "ooo-dev-tools"
VERSION = __version__

# The directory containing this file
HERE = pathlib.Path(__file__).parent
# The text of the README file
with open(HERE / "README.rst") as fh:
    README = fh.read()

setup(
    name=PKG_NAME,
    version=VERSION,
    python_requires=">=3.7.0",
    url="https://github.com/Amourspirit/python_ooo_dev_tools",
    packages=find_packages(
        exclude=[
            "src",
            "src.*",
            "env",
            "env.*",
            "cmds",
            "cmds.*",
            "*.tests",
            "*.tests.*",
            "tests.*",
            "tests",
            "*.tmp",
            "*.tmp.*",
            "tmp.*",
            "tmp",
        ]
    ),
    include_package_data=True,
    package_data={"cfg": ["cfg/*.json"]},
    author=":Barry-Thomas-Paul: Moss",
    author_email="bigbytetech@gmail.com",
    license="Apache Software License",
    install_requires=[
        'typing_extensions>=4.2.0 ;python_version<"3.8"',
        "ooouno>=0.1.15",
        "types-unopy>=0.1.7",
        "Pillow>=9.1.1",
        "lxml>=4.8.0",
    ],
    keywords=["odev", "libreoffice", "openoffice" "macro", "uno", "ooouno", "pyuno"],
    classifiers=[
        "License :: OSI Approved :: Apache Software License",
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
    description="LibreOffice Developer Tools",
    long_description_content_type="text/x-rst",
    long_description=README,
)
