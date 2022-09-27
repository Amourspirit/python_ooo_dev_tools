"""
This module is for mocking unohelper.py

DO NOT! Use this class unless you are certain what you are doing!

Read the Docs gives issues with unohelper.Base for classes that inherit from it.

The solution is to Mock unohelper during read the docs build.
"""
# it is necessary to fence using DOCS_BUILDING here.
# if not stickytape will import sphinx modules.

from .mock_g import DOCS_BUILDING
if DOCS_BUILDING:
    from sphinx.ext.autodoc.mock import _MockObject
    class Base(_MockObject):
        pass
else:
    import unohelper
    class Base(unohelper.Base):
        pass