.. _guide_lo_pip_windows_install:

Windows - Install pip packages into LibreOffice
===============================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 1

Overview
--------

Install pip_ into LibreOffice on Windows allows you to install python packages and use them in LibreOffice.

The process is essentially to install pip and then use it to install other python packages.

Sometimes you may want to install in a isolated virtual environment. In this case see :ref:`guide_windows_manual_venv`.

.. note::

    This guide assumes you have already installed LibreOffice.

    Anywhere you see ``<username>`` it needs to be replaced with your Windows username.

Install pip
-----------

LibreOffice already has python installed on Windows.
For this guide we will assume LibreOffice is installed at ``C:\Program Files\LibreOffice``.

In a PowerShell terminal navigate to ``program`` directory for you LibreOffice installation.

.. code-block:: powershell

    cd "C:\Program Files\LibreOffice\program"

.. hint::

    In window file manager if you hold down the shift key while right clicking a folder then the popup menu should include ``Open PowerShell window here``.

In PowerShell run the following command to install pip

.. code-block:: powershell

    (Invoke-WebRequest -Uri https://bootstrap.pypa.io/get-pip.py -UseBasicParsing).Content | .\python.exe -

You may get a warning that the pip install location is not on that path. This warning can be ignored.

.. code-block:: powershell

    [C:\Program Files\LibreOffice\program\]
    >(Invoke-WebRequest -Uri https://bootstrap.pypa.io/get-pip.py -UseBasicParsing).Content | .\python.exe -
    Defaulting to user installation because normal site-packages is not writeable
    Collecting pip
    Using cached pip-23.1.2-py3-none-any.whl (2.1 MB)
    Installing collected packages: pip
    WARNING: The scripts pip.exe, pip3.8.exe and pip3.exe are installed in 'C:\Users\<username>\AppData\Roaming\Python\Python38\Scripts' which is not on PATH.
    Consider adding this directory to PATH or, if you prefer to suppress this warning, use --no-warn-script-location.
    Successfully installed pip-23.1.2

.. note::

    Notice "Defaulting to user installation because normal site-packages is not writeable".
    To write for all users run PowerShell as Administrator.

Check pip version. A successful version check shows that ``pip`` is indeed on a path know to LibreOffice python.

.. code-block:: powershell

    >.\python.exe -m pip --version
    pip 23.1.2 from C:\Users\<username>\AppData\Roaming\Python\Python38\site-packages\pip (python 3.8)
    [C:\Program Files\LibreOffice\program\]

Install a python package.
We will install ooo-dev-tools_ for testing.

.. code-block:: powershell

    >.\python.exe -m pip install ooo-dev-tools
    Defaulting to user installation because normal site-packages is not writeable
    Collecting ooo-dev-tools
    Downloading ooo_dev_tools-0.11.6-py3-none-any.whl (2.2 MB)
        ---------------------------------------- 2.2/2.2 MB 4.3 MB/s eta 0:00:00
    Collecting lxml>=4.9.2 (from ooo-dev-tools)
    Using cached lxml-4.9.2-cp38-cp38-win_amd64.whl (3.9 MB)
    Collecting ooouno>=2.1.2 (from ooo-dev-tools)
    Using cached ooouno-2.1.2-py3-none-any.whl (9.8 MB)
    Collecting types-unopy>=1.2.3 (from ooouno>=2.1.2->ooo-dev-tools)
    Using cached types_unopy-1.2.3-py3-none-any.whl (5.2 MB)
    Collecting typing-extensions<5.0.0,>=4.6.2 (from ooouno>=2.1.2->ooo-dev-tools)
    Using cached typing_extensions-4.6.3-py3-none-any.whl (31 kB)
    Collecting types-uno-script>=0.1.1 (from types-unopy>=1.2.3->ooouno>=2.1.2->ooo-dev-tools)
    Using cached types_uno_script-0.1.1-py3-none-any.whl (9.3 kB)
    Installing collected packages: typing-extensions, types-uno-script, lxml, types-unopy, ooouno, ooo-dev-tools
    Successfully installed lxml-4.9.2 ooo-dev-tools-0.11.6 ooouno-2.1.2 types-uno-script-0.1.1 types-unopy-1.2.3 typing-extensions-4.6.3
    [C:\Program Files\LibreOffice\program\]

.. _guide_lo_pip_windows_install_testing_pkg:

Test installed package
----------------------

For a test we can write Hello World into a new Writer document.

With ooo-dev-tools_ installed we can now run LibreOffice python right on the command line and interact with LibreOffice.
Alternatively run a script in the APSO console as seen in :ref:`guide_lo_portable_pip_windows_install_test`.
This simple script starts python, Loads LibreOffice Writer, and writes ``Hello World!``.

.. code-block:: python

    [C:\Program Files\LibreOffice\program\]
    >.\python.exe
    Python 3.8.16 (default, Apr 28 2023, 02:01:33) [MSC v.1929 64 bit (AMD64)] on win32
    Type "help", "copyright", "credits" or "license" for more information.
    >>> from ooodev.utils.lo import Lo
    >>> from ooodev.office.write import Write
    >>> from ooodev.utils.gui import GUI
    >>> 
    >>> def say_hello():
    ...     cursor = Write.get_cursor(Write.active_doc)
    ...     Write.append_para(cursor=cursor, text="Hello World!")
    ...
    >>> _ = Lo.load_office(Lo.ConnectSocket())
    >>> doc = Write.create_doc()
    >>> GUI.set_visible(visible=True, doc=doc)
    >>> say_hello()
    >>> Lo.close_doc(doc)
    >>> Lo.close_office()
    True
    >>>

The resulting document should look like :numref:`b370cae2-a6f6-41b7-9dfb-be6e4514bbf6`

.. cssclass:: screen_shot

    .. _b370cae2-a6f6-41b7-9dfb-be6e4514bbf6:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/b370cae2-a6f6-41b7-9dfb-be6e4514bbf6
        :alt: LibreOffice Writer Hello World
        :figclass: align-center

        LibreOffice Writer Hello World

Recommended Python Packages
---------------------------

- ooo-dev-tools_ is a Python package that provides a framework to help with development of LibreOffice python projects. See |odev_docs|_.
- types-scriptforge_ is a Python package that provides type hints for the ScriptForge_ library.
- types-unopy_ is a Python package the has typings for the full LibreOffice API

.. note::

    Both ooo-dev-tools_ and types-scriptforge_ install the types-unopy_ package.


Related Links
-------------

- :ref:`guide_apso_installation`
- :ref:`guide_lo_portable_pip_windows_install`
- :ref:`guide_windows_manual_venv`
- :ref:`guide_windows_poetry_venv`
- |win_pre_venv|_

.. _ooo-dev-tools: https://pypi.org/project/ooo-dev-tools/
.. _pip: https://pip.pypa.io/en/stable/

.. |win_pre_venv| replace:: Pre-configured virtual environments for Windows
.. _win_pre_venv: https://github.com/Amourspirit/lo-support_file/tree/main/virtual_environments/windows

.. |odev_docs| replace:: OooDev Docs
.. _odev_docs: https://python-ooo-dev-tools.readthedocs.io/en/latest/index.html
.. _types-scriptforge: https://pypi.org/project/types-scriptforge/
.. _scriptforge: https://gitlab.com/LibreOfficiant/scriptforge
.. _types-unopy: https://pypi.org/project/types-unopy/