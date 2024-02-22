# coding: utf-8
"""
Module with globals for dealing with documentation building.

DO NOT! Use unless you are certain what you are doing!
"""
import os

ON_RTD = os.environ.get("READTHEDOCS", None) == "True"
"""True if environment is running on read the docs; Otherwise, False. https://readthedocs.org"""
LOCAL_DOCS_BUILDING = os.environ.get("DOCS_BUILDING", None) == "True"
"""True if DOCS_BUILDING environment var is set to 'True' in docs.py"""
DOCS_BUILDING = ON_RTD or LOCAL_DOCS_BUILDING
"""True if LOCAL_DOCS_BUILDING or ON_RTD is True; Otherwise False"""

FULL_IMPORT = os.environ.get("OOO_SCRIPT_COMPILING", None) == "True"
"""
This is True if the OOO-SCRIPT is compiling the script and all late imports should be imported while this set; Otherwise, False.

.. versionadded:: 0.30.0
"""
