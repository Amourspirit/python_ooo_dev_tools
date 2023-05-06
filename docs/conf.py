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

odevgui_win_url = "https://ooo-dev-tools-gui-win.readthedocs.io/en/latest/"

# The full version, including alpha/beta/rc tags
release = __version__

# region Configuration methods


def get_active_branch_name():
    head_dir = _ROOT_PATH / ".git" / "HEAD"
    with head_dir.open("r") as f:
        content = f.read().splitlines()

    for line in content:
        if line[:4] == "ref:":
            return line.partition("refs/heads/")[2]


def get_spell_dictionaries() -> list:
    p = _DOCS_PATH.absolute().resolve() / "internal" / "dict"
    dict_gen = p.glob("spelling_*.*")
    return [str(d) for d in dict_gen if d.is_file()]


# endregion Configuration methods


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
    "sphinx_toolbox.more_autodoc.overloads",
    "sphinx.ext.napoleon",
    "sphinx.ext.todo",
    "sphinx.ext.extlinks",
    "sphinx.ext.intersphinx",
    "sphinx_autodoc_typehints",
    "sphinx_tabs.tabs",
    "sphinx_design",
    "sphinxcontrib.spelling",
    "sphinx-prompt",
    "sphinx_substitution_extensions",
]
# "sphinx.ext.linkcode",
# sphinx_tabs.tabs docs: https://sphinx-tabs.readthedocs.io/en/latest/

# "sphinx_design"
# https://sphinx-design.readthedocs.io/en/rtd-theme/get_started.html

# region spelling
# https://sphinxcontrib-spelling.readthedocs.io/en/latest/

# sphinx_substitution_extensions
# https://github.com/adamtheturtle/sphinx-substitution-extensions

# https://www.sphinx-doc.org/en/master/usage/configuration.html#confval-suppress_warnings
# https://github.com/sphinx-doc/sphinx/issues/4961
# List of zero or more Sphinx-specific warning categories to be squelched (i.e.,
# suppressed, ignored).
suppress_warnings = [
    # FIXME: *THIS IS TERRIBLE.* Generally speaking, we do want Sphinx to inform
    # us about cross-referencing failures. Remove this hack entirely after Sphinx
    # resolves this open issue:
    #    https://github.com/sphinx-doc/sphinx/issues/4961
    # Squelch mostly ignorable warnings resembling:
    #     WARNING: more than one target found for cross-reference 'TypeHint':
    #     beartype.door._doorcls.TypeHint, beartype.door.TypeHint
    "ref.python",
]

spelling_word_list_filename = get_spell_dictionaries()

spelling_show_suggestions = True
spelling_ignore_pypi_package_names = True
spelling_ignore_contributor_names = True
spelling_ignore_acronyms = True

# https://sphinxcontrib-spelling.readthedocs.io/en/latest/customize.html
spelling_exclude_patterns = [".venv/"]

# spell checking;
#   run sphinx-build -b spelling . _build
#       this will check for any spelling and create folders with *.spelling files if there are errors.
#       open each *.spelling file and find any spelling errors and fix them in corresponding files.
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

html_css_files = ["https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.2.1/css/all.min.css"]
if html_theme == "sphinx_rtd_theme":
    html_css_files.append("css/readthedocs_custom.css")

if "sphinx_rtd_dark_mode" in extensions:
    html_css_files.append("css/readthedocs_dark.css")

html_css_files.append("css/style_custom.css")

html_js_files = [
    "js/custom.js",
]

# Napoleon settings
# https://www.sphinx-doc.org/en/master/usage/extensions/napoleon.html
napoleon_google_docstring = True
napoleon_include_init_with_doc = True

# https://fossies.org/linux/Sphinx/doc/usage/extensions/autodoc.rst
# This value controls how to represent type hints. The setting takes the following values:
autodoc_typehints = "description"

# https://documentation.help/Sphinx/autodoc.html
autoclass_content = "init"


# see: https://www.sphinx-doc.org/en/master/usage/extensions/autodoc.html#confval-autodoc_mock_imports
# see: https://read-the-docs.readthedocs.io/en/latest/faq.html#i-get-import-errors-on-libraries-that-depend-on-c-modules
# on read the docs I was getting errors WARNING: autodoc: failed to import class - No module named 'uno'
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
    "XChartDocument": "XChartDocument",
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


autodoc_typehints_format = "short"


# https://stackoverflow.com/questions/9698702/how-do-i-create-a-global-role-roles-in-sphinx
# custom global roles or any other rst to include

# sphinx includes s5defs.txt that has baked in roles but must be included.
# style_custom.css contains the colors that match these roles
# https://stackoverflow.com/questions/3702865/sphinx-restructuredtext-set-color-for-a-single-word/60991308#60991308

rst_prolog_lst = [
    ".. include:: <s5defs.txt>",
    "",
    ".. role:: event(doc)",
    "",
    ".. role:: eventref(ref)",
    "",
    ".. |app_name| replace:: OOO Development Tools",
    "",
    ".. |app_name_bold| replace:: **OOO Development Tools**",
    "",
    ".. |odev| replace:: OooDev",
    "",
    ".. |ooouno| replace:: ooouno library",
    ".. _ooouno: https://pypi.org/project/ooouno/",
    "",
    ".. |odevgui_win| replace:: OooDev GUI Automation for windows",
    f".. _odevgui_win: {odevgui_win_url}",
    "",
]
rst_prolog_lst.append(f".. |app_ver| replace:: {__version__}\n")

# add extra roles for custom theme colors.
# unlike the colors style_custom.css, these colors can be change by
# changing colors css vars that start with -t-color-
with open("roles/theme_color_roles.txt", "r") as file:
    rst_prolog = file.read()

rst_prolog += "\n" + "\n".join(rst_prolog_lst)

# set if figures can be referenced as numbers. Default is False
# https://www.sphinx-doc.org/en/master/usage/configuration.html#confval-numfig
numfig = True
# numfig_secnum_depth = 1

# set is todo's will show up.
# a master list of todo's will be on bottom of main page.
# https://www.sphinx-doc.org/en/master/usage/extensions/todo.html#module-sphinx.ext.todo
todo_include_todos = False

# region sphinx.ext.extlinks – Markup to shorten external links
# https://documentation.help/Sphinx/extlinks.html
odev_src_root = "../../_modules/ooodev"
extlinks = {
    "odev_src_calc_meth": (odev_src_root + "/office/calc.html#Calc.%s", "Calc.%s()"),
    "odev_src_draw_meth": (odev_src_root + "/office/draw.html#Draw.%s", "Draw.%s()"),
    "odev_src_write_meth": (odev_src_root + "/office/write.html#Write.%s", "Write.%s()"),
    "odev_src_gui_meth": (odev_src_root + "/utils/gui.html#GUI.%s", "GUI.%s()"),
    "odev_src_lo_meth": (odev_src_root + "/utils/lo.html#Lo.%s", "Lo.%s()"),
    "odev_src_info_meth": (odev_src_root + "/utils/info.html#Info.%s", "Info.%s()"),
    "odev_src_props_meth": (odev_src_root + "/utils/props.html#Props.%s", "Props.%s()"),
    "odev_src_fileio_meth": (odev_src_root + "/utils/file_io.html#FileIO.%s", "FileIO.%s()"),
    "odev_src_loinst_meth": (odev_src_root + "/utils/inst/lo/lo_inst.html#LoInst.%s", "LoInst.%s()"),
}
# endregion sphinx.ext.extlinks – Markup to shorten external links

# region intersphinx
intersphinx_mapping = {"odevguiwin": (odevgui_win_url, None)}

# endregion intersphinx
