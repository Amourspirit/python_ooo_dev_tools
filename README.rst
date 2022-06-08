OOO Development Tools
---------------------

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


Why Python?
+++++++++++

Macros are pieces of programming code that runs in office suites and helps automate routine tasks.
Specifically, in LibreOffice API these codes can be written with so many programming languages thanks
to the Universal Network Objects (UNO). Among them are: Open/LibreOffice Basic (Thanks to Foad S Farimani for the correction :) ), Java, C/C++, Javascript, Python.

So which language should we use? Since LibreOffice is multi-platform we can use our documents at different
platforms like Mac, Windows, and Linux. So we need a cross-platform language to run our macros at different
platforms. We can eliminate Visual Basic because of that.

Java and C/C++ require compilation, are much more complex and verbose. So we can eliminate these too.

Probably we will have some problems while working with numbers if we choose JavaScript.
For example it has rounding errors ( 0.1 + 0.2 does not equals 0.3 in Javascript).
So we can eliminate this too.
But Python is very powerful at numeric computation thanks to its libraries.
Libraries like Numpy and Numexpr is excellent for this job.
So we should choose Python 3 for macro programming [Ref1]_.

Documentation
+++++++++++++

Read `documentation <https://python-ooo-dev-tools.readthedocs.io/en/latest/>`_

Installation
++++++++++++

PIP
***

**ooo-dev-tools** `PyPI <https://pypi.org/project/ooo-dev-tools/>`_

.. code-block:: bash

    $ pip install ooo-dev-tools


Modules
+++++++

Currently the ``Calc`` Module has been fully tested.

The ``Write`` Module is also in this release but has very limited testing at this point.

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

Include modules that have not yet been fully tested:
    - Color (Various color utils)
    - DateUtil (Date Time utilities)
    - FileIO (File Input and Output for working with LO)
    - GUI (Various Gui methods for manipulating LO Windows)
    - Images (Various methods for working with Images)
    - Lo (Various methods common to LO applications)
    - Props (Various methods for working with the many API properties)
    - XML (Various methods for working with LO xml data)


Release Info
++++++++++++

This is a beta release. Calc is the first module due to its popularity.

Inspiration
+++++++++++

Much of this project is inspired by the work of Dr. Andrew Davison
and the work on `Java LibreOffice Programming <http://fivedots.coe.psu.ac.th/~ad/jlop>`_

See `LibreOffice Programming <https://flywire.github.io/lo-p/>`_ that aims to gradually explain this content in a python context.


Other
+++++

**Figure 1:** Calc Automation

.. figure:: https://user-images.githubusercontent.com/4193389/172459702-26f87b92-6986-4d8f-b627-0c5e8602b3c5.gif
   :alt: Calc automation example gif.


Related projects
++++++++++++++++

LibreOffice API Typings

 * `LibreOffice API Typings <https://github.com/Amourspirit/python-types-unopy>`_
 * `ScriptForge Typings <https://github.com/Amourspirit/python-types-scriptforge>`_
 * `Access2base Typings <https://github.com/Amourspirit/python-types-access2base>`_
 * `LibreOffice UNO Typings <https://github.com/Amourspirit/python-types-uno-script>`_
 * `LibreOffice Developer Search <https://github.com/Amourspirit/python_lo_dev_search>`_
 * `LibreOffice Python UNO Examples <https://github.com/Amourspirit/python-ooouno-ex>`_
 * `OOOUNO <https://github.com/Amourspirit/python-ooouno>`_
 * `OOO UNO TEMPLATE <https://github.com/Amourspirit/ooo_uno_tmpl>`_

.. [Ref1] `Macro Programming in OpenOffice/LibreOffice with using Python <https://medium.com/analytics-vidhya/macro-programming-in-openoffice-libreoffice-with-using-python-en-a37465e9bfa5>`_

.. _LibreOffice: http://www.libreoffice.org/

.. |lic| image:: https://img.shields.io/github/license/Amourspirit/python_ooo_dev_tools
    :alt: License Apache

.. |pver| image:: https://img.shields.io/pypi/pyversions/python_ooo_dev_tools
    :alt: PyPI - Python Version

.. |pwheel| image:: https://img.shields.io/pypi/wheel/python_ooo_dev_tools
    :alt: PyPI - Wheel

.. |github| image:: https://img.shields.io/badge/GitHub-100000?style=plastic&logo=github&logoColor=white
    :target: https://github.com/Amourspirit/python_ooo_dev_tools
    :alt: Github