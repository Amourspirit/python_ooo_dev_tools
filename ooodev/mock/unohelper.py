"""
This module is for mocking unohelper.py

DO NOT! Use this class unless you are certain what you are doing!

Read the Docs gives issues with unohelper.Base for classes that inherit from it.

The solution is to Mock unohelper during read the docs build
"""
from sphinx.ext.autodoc.mock import _MockObject
class Base(_MockObject):
    pass