# Configuration file for the Sphinx documentation builder.
#
# This file only contains a selection of the most common options. For a full
# list see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Path setup --------------------------------------------------------------

# If extensions (or modules to document with autodoc) are in another directory,
# add these directories to sys.path here. If the directory is relative to the
# documentation root, use os.path.abspath to make it absolute, like shown here.
#
import os
import sys

sys.path.insert(0, os.path.abspath(".."))
from ooodev import __version__

os.environ["DOCS_BUILDING"] = "True"
os.environ["ooouno_ignore_runtime"] = "True"
# os.environ["ooouno_ignore_import_error"] = "True"

# -- Project information -----------------------------------------------------

project = "OOO Development Tools"
copyright = "2022, :Barry-Thomas-Paul: Moss"
author = ":Barry-Thomas-Paul: Moss"

# The full version, including alpha/beta/rc tags
release = __version__


# -- General configuration ---------------------------------------------------

# Add any Sphinx extension module names here, as strings. They can be
# extensions coming with Sphinx (named 'sphinx.ext.*') or your custom ones.
#
# To use sphinx.ext.napoleon with sphinx-autodoc-typehints, make sure you load sphinx.ext.napoleon first,
# before sphinx-autodoc-typehints.
extensions = [
    "sphinx_rtd_theme",
    "sphinx_rtd_dark_mode",
    "sphinx.ext.autodoc",
    "sphinx_toolbox.collapse",
    "sphinx_toolbox.more_autodoc.autonamedtuple",
    "sphinx_toolbox.more_autodoc.typevars",
    "sphinx_toolbox.more_autodoc.autoprotocol",
    "sphinx.ext.napoleon",
    "sphinx_autodoc_typehints",
]
# Add any paths that contain templates here, relative to this directory.
templates_path = ["_templates"]

# List of patterns, relative to source directory, that match files and
# directories to ignore when looking for source files.
# This pattern also affects html_static_path and html_extra_path.
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store"]


# -- Options for HTML output -------------------------------------------------

# The theme to use for HTML and HTML Help pages.  See the documentation for
# a list of builtin themes.
#
# html_theme = 'alabaster'
html_theme = "sphinx_rtd_theme"

# Add any paths that contain custom static files (such as style sheets) here,
# relative to this directory. They are copied after the builtin static files,
# so a file named "default.css" will overwrite the builtin "default.css".
html_static_path = ["_static"]

html_css_files = ["https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.1.1/css/all.min.css"]
html_css_files.append("css/readthedocs_custom.css")
if html_theme == "sphinx_rtd_theme":
    html_css_files.append("css/readthedocs_dark.css")

# Napoleon settings
# https://www.sphinx-doc.org/en/master/usage/extensions/napoleon.html
napoleon_google_docstring = True
napoleon_include_init_with_doc = True

# https://fossies.org/linux/Sphinx/doc/usage/extensions/autodoc.rst
# This value controls how to represent typehints. The setting takes the following values:
autodoc_typehints = 'description'


# see: https://www.sphinx-doc.org/en/master/usage/extensions/autodoc.html#confval-autodoc_mock_imports
# see: https://read-the-docs.readthedocs.io/en/latest/faq.html#i-get-import-errors-on-libraries-that-depend-on-c-modules
# on read the docs I was getting errros WARNING: autodoc: failed to import class - No module named 'uno'
# solution autodoc_mock_imports, for some reason after adding uno, unohelper I also had to include com.
# com.sun.star.__init__.py raises an Import error by design.
autodoc_mock_imports = ["uno", "unohelper", "lxml", "com"]
# also note that object can be mocked. See ooodev/utils/uno_enum.py

# autodoc_type_aliases
# The key of each entry is the type hint as written in the source,
# the value is how it will be written in the generated documentation.
autodoc_type_aliases = {
    "Row": "Row",
    "Column": "Column",
    "Table": "Table",
    "TupleArray": "TupleArray",
    "FloatList": "FloatList",
    "FloatTable": "FloatTable",
    "DocOrCursor": "DocOrCursor",
    "DocOrText": "DocOrText",
    "PathOrStr": "PathOrStr",
    "DictRow": "DictRow",   
    # Uno Types
    "CellAddress": "CellAddress",
    "CellRangeAddress": "CellRangeAddress",
    "PropertyValue": "PropertyValue",
    "XCell": "XCell",
    "XCellRange": "XCellRange",
    "XFrame": "XFrame",
    "XModel": "XModel",
    "XSpreadsheet": "XSpreadsheet",
    "XSpreadsheetDocument": "XSpreadsheetDocument",
    "XTextCursor": "XTextCursor",
    "XTextDocument": "XTextDocument",
    "XTextRange": "XTextRange",
}


autodoc_typehints_format = 'short'

# https://stackoverflow.com/questions/9698702/how-do-i-create-a-global-role-roles-in-sphinx
# custom global roles or any other rst to include
rst_prolog = """
.. role:: event(doc)

.. role:: eventref(ref)
"""

