OOO Development Tools
=====================

|lic| |pver| |pwheel| |test_badge|_ |github|

OOO Development Tools (OooDev) is intended for programmers who want to learn and use the
Python version of the `LibreOffice`_ API.

This allows Python to control and manipulate LibreOffice's text, drawing, presentation, spreadsheet, and database applications,
and a lot more (e.g. its spell checker, forms designer, and charting tools).

One of the aims is to develop utility code to help flatten the steep learning curve for the API.
For example, The Lo class simplifies the steps needed to initialize the API
(by creating a connection to a LibreOffice process), to open/create a document, save it,
and close down LibreOffice.

Currently this project has been tested on LibreOffice in Windows and Linux (Ubuntu).

Advantages of Python
--------------------

Macros are pieces of programming code that runs in office suites and helps automate routine tasks.
Specifically, in LibreOffice API these codes can be written with so many programming languages thanks
to the Universal Network Objects (UNO).

Since LibreOffice is multi-platform we can use our documents at different
platforms like Mac, Windows, and Linux. So we need a cross-platform language to run our macros at different
platforms.

Python has the advantage that it is cross-platform and can run inside the office environment as macros and outside
office environment on the command line.

Python has a vast set `libraries <https://pypi.org/>`_ that can be used in a project, including `Numpy <https://numpy.org/>`_ and
`Numexpr <https://github.com/pydata/numexpr>`_ which are excellent and powerful at numeric computation.

This makes Python and excellent choice with maximum flexibility.


Documentation
-------------

Docs
^^^^

Read `documentation <https://python-ooo-dev-tools.readthedocs.io/en/latest/>`_

Command Line Help
^^^^^^^^^^^^^^^^^

There are many classes and methods in this project.
For this reason OooDev has a command line tool |cli_hlp|_.
that can be used to search the documentation for classes and methods this project.
Choosing a number from a search result opens you web browser to that class or method in the documentation.

|cli_hlp|_ is built from `Sphinx CLI Help <https://github.com/Amourspirit/python-sphinx-cli-help>`__, so see the `Wiki Searching help <https://github.com/Amourspirit/python-sphinx-cli-help/wiki/Searching>`__ and substitute ``cli-hlp`` with ``odh`` on command line instructions.

Example Usage:

.. code-block:: bash

    odh hlp -s Write.append

    Choose an option (default 1):
    [0],  Cancel (or press q followed by enter)
    [1],  ooodev.office.write.Write.append                                 - method     - py
    [2],  ooodev.office.write.Write.append_date_time                       - method     - py
    [3],  ooodev.office.write.Write.append_line                            - method     - py
    [4],  ooodev.office.write.Write.append_para                            - method     - py


Note that |cli_hlp|_ is a separate project and is not required to use this project.
Also |cli_hlp|_ use python built in ``sqlite`` which is not shipped with LibreOffices's python on Windows.
This means some configurations will not allow |cli_hlp|_ to run on Windows. In this case you can install |cli_hlp|_ on a global scope and use it that way if needed.

See Also: `LibreOffice Developer Search <https://pypi.org/project/lo-dev-search/>`__

Installation
------------

EXTENSION
^^^^^^^^^

This project is also available as an extension for LibreOffice.
The OOO Development Tools extension is available on `GitHub <https://github.com/Amourspirit/libreoffice_ooodev_ext>`__ and `LibreOffice Extensions <https://extensions.libreoffice.org/en/extensions/show/41700>`__.

See the `Guide <https://python-ooo-dev-tools.readthedocs.io/en/main/guide/guide_ooodev_oxt.html>`__ for installation and usage.

PIP
^^^

**ooo-dev-tools** `PyPI <https://pypi.org/project/ooo-dev-tools/>`_

.. code-block:: bash

    pip install ooo-dev-tools

Note: Support for python ``3.7`` was dropped in version ``0.10.0``

Modules
-------

Currently there are more than ``4,000`` classes in this framework.

Include modules:
    - Calc (Calc)
    - Write (Write)
    - Draw (LibreOffice Draw/Impress)
    - Forms (Support for building forms)
    - Dialogs (Build dialog forms)
    - GUI (Various GUI methods for manipulating LO Windows)
    - Lo (Various methods common to LO applications)
    - FileIO (File Input and Output for working with LO)
    - Format (Format Module — **hundreds of classes** — for Styling and modifying the many Documents and Sheets properties.)
    - Props (Various methods setting and getting the many properties of Office objects)
    - Info (Various method for getting information about LO applications)
    - Color (Various color utils)
    - DateUtil (Date Time utilities)
    - ImagesLo (Various methods for working with Images)
    - Props (Various methods for working with the many API properties)
    - Chart2 (charting)
    - Chart (charting)
    - Gallery (Methods for accessing and reporting on the Gallery)
    - Theme (Access to LibreOffice Theme Properties)
    - Units (Various unit methods and classes for passing different kinds of units in LibreOffice such as inches, millimeters, points, pixels.)
    - And more ...

Future releases will add:
    - Base (LibreOffice Base)
    - Clip (clipboard support)
    - Mail (Mail service provider)
    - Print (Print service provider)
    - And more ...

Inspiration
-----------

Much of this project is inspired by the work of Dr. Andrew Davison.
An archive archive of the Java code is available at `GitHub - LibreOffice Java Programming <https://github.com/Amourspirit/libreoffice_lop_java>`__.

See Also:

- `LibreOffice Programming <https://flywire.github.io/lo-p/>`_ that aims to gradually explain this content in a python context.
- `Python LibreOffice Programming - Preface <https://python-ooo-dev-tools.readthedocs.io/en/latest/odev/preface.html>`__


Other
-----

**Figure 1:** Calc Find and Replace Automation Example

.. figure:: https://user-images.githubusercontent.com/4193389/172609472-536a94de-9bf6-4668-ac9f-a55f12dfc817.gif
    :alt: Calc Find and Replace Automation


Related projects
----------------

LibreOffice API Typing's

 * `LibreOffice API Typings <https://github.com/Amourspirit/python-types-unopy>`_
 * `ScriptForge Typings <https://github.com/Amourspirit/python-types-scriptforge>`_
 * `Access2base Typings <https://github.com/Amourspirit/python-types-access2base>`_
 * `LibreOffice UNO Typings <https://github.com/Amourspirit/python-types-uno-script>`_
 * `LibreOffice Developer Search <https://github.com/Amourspirit/python_lo_dev_search>`_
 * `LibreOffice Python UNO Examples <https://github.com/Amourspirit/python-ooouno-ex>`_
 * `OOOUNO Project <https://github.com/Amourspirit/python-ooouno>`_
 * `OOO UNO TEMPLATE <https://github.com/Amourspirit/ooo_uno_tmpl>`_

.. _LibreOffice: http://www.libreoffice.org/

.. |lic| image:: https://img.shields.io/github/license/Amourspirit/python_ooo_dev_tools
    :alt: License Apache

.. |pver| image:: https://img.shields.io/pypi/pyversions/ooo-dev-tools
    :alt: PyPI - Python Version

.. |pwheel| image:: https://img.shields.io/pypi/wheel/ooo-dev-tools
    :alt: PyPI - Wheel

.. |github| image:: https://img.shields.io/badge/GitHub-100000?style=plastic&logo=github&logoColor=white
    :target: https://github.com/Amourspirit/python_ooo_dev_tools
    :alt: Github

.. |test_badge| image:: https://github.com/Amourspirit/python_ooo_dev_tools/actions/workflows/python-app-test.yml/badge.svg
    :alt: Test Badge

.. _test_badge: https://github.com/Amourspirit/python_ooo_dev_tools/actions/workflows/python-app-test.yml

.. |cli_hlp| replace:: OooDev CLI Help
.. _cli_hlp: https://github.com/Amourspirit/python-ooodev-cli-hlp#readme