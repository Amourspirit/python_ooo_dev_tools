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
from pathlib import Path

_DOCS_PATH = Path(__file__).parent
_ROOT_PATH = _DOCS_PATH.parent

sys.path.insert(0, str(_ROOT_PATH))
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
    "sphinx.ext.viewcode",
    "sphinx_toolbox.collapse",
    "sphinx_toolbox.more_autodoc.autonamedtuple",
    "sphinx_toolbox.more_autodoc.typevars",
    "sphinx_toolbox.more_autodoc.autoprotocol",
    "sphinx.ext.napoleon",
    "sphinx.ext.todo",
    "sphinx_autodoc_typehints",
    "sphinx_tabs.tabs",
    "sphinxcontrib.spelling",
]
    # "sphinx.ext.linkcode",
    # sphinx_tabs.tabs docs: https://sphinx-tabs.readthedocs.io/en/latest/

# region spelling
# https://sphinxcontrib-spelling.readthedocs.io/en/latest/


def get_spell_dictionaries() -> list:

    p = _DOCS_PATH.absolute().resolve() / "internal" / "dict"
    dict_gen = p.glob('spelling_*.*')
    return [str(d) for d in dict_gen if d.is_file()]

spelling_word_list_filename = get_spell_dictionaries()

spelling_show_suggestions = True
spelling_ignore_pypi_package_names = True
spelling_ignore_contributor_names = True
spelling_ignore_acronyms=True

# spell checking;
#   run sphinx-build -b spelling . _build
#       this will checkfor any spelling and create folders with *.spelling files if there are errors.
#       open each *.spelling file and find any spelling errors and fix them in corrsponding files.
#
# spelling_book.txt contains all spelling exceptions related to book in /docs/odev
# spelling_code.txt contains all spelling exceptions related to python doc strings.

# endregion spelling

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

html_js_files = [
    'js/custom.js',
]

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
    "Column": "Column",
    "DictRow": "DictRow",
    "DocOrCursor": "DocOrCursor",
    "DocOrText": "DocOrText",
    "FloatList": "FloatList",
    "FloatTable": "FloatTable",
    "PathOrStr": "PathOrStr",
    "Row": "Row",
    "Table": "Table",
    "TupleArray": "TupleArray",
    # Uno Types
    "CellAddress": "CellAddress",
    "CellRangeAddress": "CellRangeAddress",
    "Locale": "Locale",
    "PropertyValue": "PropertyValue",
    "XCell": "XCell",
    "XCellRange": "XCellRange",
    "XComponent": "XComponent",
    "XDrawPage": "XDrawPage",
    "XDrawPages": "XDrawPages",
    "XDrawView": "XDrawView",
    "XFrame": "XFrame",
    "XModel": "XModel",
    "XNameContainer": "XNameContainer",
    "XShape": "XShape",
    "XSpreadsheet": "XSpreadsheet",
    "XSpreadsheetDocument": "XSpreadsheetDocument",
    "XTextCursor": "XTextCursor",
    "XTextDocument": "XTextDocument",
    "XTextRange": "XTextRange",
}


autodoc_typehints_format = 'short'


# https://stackoverflow.com/questions/9698702/how-do-i-create-a-global-role-roles-in-sphinx
# custom global roles or any other rst to include

rst_prolog_lst = []
rst_prolog_lst.append(".. role:: event(doc)")
rst_prolog_lst.append("")
rst_prolog_lst.append(".. role:: eventref(ref)")
rst_prolog_lst.append("")
rst_prolog_lst.append(".. |app_name| replace:: OOO Development Tools")
rst_prolog_lst.append("")
rst_prolog_lst.append(".. |app_name_bold| replace:: **OOO Development Tools**")
rst_prolog_lst.append("")
rst_prolog_lst.append(".. |odev| replace:: ODEV")
rst_prolog_lst.append("")
rst_prolog_lst.append(".. _ooouno: https://pypi.org/project/ooouno/")
rst_prolog_lst.append("")
rst_prolog_lst.append(f".. |app_ver| replace:: {__version__}")


rst_prolog = "\n".join(rst_prolog_lst)

# set if figures can be referenced as numers. Defalut is False
numfig = True

# set is todo's will show up.
# a master list of todo's will be on bottom of main page.
# https://www.sphinx-doc.org/en/master/usage/extensions/todo.html#module-sphinx.ext.todo
todo_include_todos = False

# region Not currently Used

# Add external links to source code

def get_active_branch_name():

    head_dir = _ROOT_PATH / ".git" / "HEAD"
    with head_dir.open("r") as f: content = f.read().splitlines()

    for line in content:
        if line[0:4] == "ref:":
            return line.partition("refs/heads/")[2]

# endregion Not currently Used