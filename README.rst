OOO Development Tools
=====================

|lic| |pver| |pwheel| |github|

OOO Development Tools (ODEV) is intended for programmers who want to learn and use the
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

Read `documentation <https://python-ooo-dev-tools.readthedocs.io/en/latest/>`_


Installation
------------

PIP
^^^

**ooo-dev-tools** `PyPI <https://pypi.org/project/ooo-dev-tools/>`_

.. code-block:: bash

    $ pip install ooo-dev-tools


Modules
-------

Currently the ``Calc`` Module and the ``Write`` module are released.


Future releases will add:
    - Base (LibreOffice Base)
    - Chart (charting)
    - Chart2 (charging)
    - Clip (clipboard support)
    - Dialogs (Build dialog forms)
    - Draw (LibreOffice Draw)
    - Forms (Support for building forms)
    - Gallery (Methods for accessing and reporting on the Gallery)
    - Mail (Mail service provider)
    - Print (Print service provider)
    - And more ...

Include modules still in beta:
    - Color (Various color utils)
    - DateUtil (Date Time utilities)
    - FileIO (File Input and Output for working with LO)
    - GUI (Various GUI methods for manipulating LO Windows)
    - ImagesLo (Various methods for working with Images)
    - Lo (Various methods common to LO applications)
    - Props (Various methods for working with the many API properties)



Inspiration
-----------

Much of this project is inspired by the work of Dr. Andrew Davison
and the work on `Java LibreOffice Programming <http://fivedots.coe.psu.ac.th/~ad/jlop>`_

See `LibreOffice Programming <https://flywire.github.io/lo-p/>`_ that aims to gradually explain this content in a python context.


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