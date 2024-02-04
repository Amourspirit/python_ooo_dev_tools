.. _guide_windows_poetry_venv:

Windows - Creating a Poetry Virtual Environment for LibreOffice
===============================================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

LibreOffice on Windows is shipped with its own embedded Python.
Unfortunately, this Python is not compatible with the Python Poetry_ needs to use.

To work around this, we need to create a virtual environment with a compatible Python version.
This virtual environment will need to be able to be used by both Poetry and LibreOffice.
This is not straightforward, as LibreOffice does not support virtual environments unless it is configured in a specific way.

If you only need pip_ and not Poetry_, then see :ref:`guide_windows_manual_venv` to create a virtual environment.

This guide will show you how to create a virtual environment with Poetry_ and pyenv_.

The reason pyenv_ is used is because it provides a way to install a compatible Python version on Windows that matches
the version of Python shipped with LibreOffice.

See :ref:`guide_configure_poetry` for poetry installation instructions.

.. note::

    This guide assumes you have already installed LibreOffice.

    Anywhere you see ``<username>`` it needs to be replaced with your Windows username.

Installing pyenv
----------------

To install pyenv_ follow the instructions on the pyenv_ GitHub page.

How to use pyenv_ however, is out of the scope of this guide, so only limited instructions will be provided.

Installing a python version
---------------------------

First, we will need to get the version of Python that LibreOffice uses.

.. code-block:: powershell

    &"C:\Program Files\LibreOffice\program\python.exe" --version

This will output something like: ``3.8.16``

List the available versions of Python that can be installed with pyenv_.
Limit the output to only versions that match the major and minor version of the Python version that LibreOffice uses.

.. code-block:: powershell

    pyenv install -l | findstr 3.8

At the time of writing this guide pyenv_ does not have a ``3.8.16`` version of Python but it does have a ``3.8.10`` version.
As long and the major and minor versions match, it will be fine. So any ``3.8.x`` version will work.

Install the ``3.8.10`` version of Python:

.. code-block:: powershell

    pyenv install 3.8.10

pyenv_ will download and install the Python version.
The installation path will be something like: ``C:\Users\<username>\.pyenv\pyenv-win\versions\3.8.10``

Creating a virtual environment
------------------------------

Now that we have a compatible version of Python installed, we can create a virtual environment.
For the purpose of this guide, we will create a virtual environment in the ``D:\tmp\project`` directory.

.. code-block:: powershell

    cd D:\tmp\project

Create the virtual environment with the ``3.8.10`` version of Python:

.. code-block:: powershell

    &"C:\Users\<username>\.pyenv\pyenv-win\versions\3.8.10\python.exe" -m venv --without-pip .venv

.. note::

    The ``--without-pip`` option is used because we will be using Poetry_ to manage the packages. And pip_ is not needed.
    If you need pip_ for some reason then you can omit the ``--without-pip`` option.

Activate the virtual environment:

.. code-block:: powershell

    .\.venv\Scripts\Activate.ps1

Check the version of Python:

.. code-block:: powershell

    (.venv) PS D:\tmp\project> python --version
    Python 3.8.10

Initialize poetry project
-------------------------

Use poetry_ to initialize a project. We will add packages later as there are some additional steps that need to be done.

.. code-block:: powershell

    cd D:\tmp\project
    poetry init

Output:

.. code-block:: text

    This command will guide you through creating your pyproject.toml config.

    Package name [project]:
    Version [0.1.0]:
    Description []:  My fantastic project
    Author [Secret Name <secret@name.nowhere>, n to skip]:
    License []:  MIT
    Compatible Python versions [^3.9]:  ^3.8

    Would you like to define your main dependencies interactively? (yes/no) [yes] n
    Would you like to define your development dependencies interactively? (yes/no) [yes] n

The generated ``pyproject.toml`` file will look something like:

.. code-block:: toml

    [tool.poetry]
    name = "project"
    version = "0.1.0"
    description = "My fantastic project"
    authors = ["Secret Name <secret@name.nowhere>"]
    license = "MIT"
    readme = "README.md"

    [tool.poetry.dependencies]
    python = "^3.8"

    [build-system]
    requires = ["poetry-core"]
    build-backend = "poetry.core.masonry.api"

Install OOOENV
--------------

oooenv_ is a Python package that allows you to auto configure a virtual environment to be used by LibreOffice.

Install oooenv_ in the virtual environment:

.. code-block:: powershell

    poetry add oooenv --group=dev

.. note::

    The ``--group=dev`` option is used because we only need oooenv_ for development purposes.
    This option instructs Poetry_ to only add oooenv_ to the ``dev-dependencies`` section of the ``pyproject.toml`` file.

Do a version check to make sure it is installed:

.. code-block:: powershell

    (.venv) PS D:\tmp\project> oooenv --version
    0.2.0

Toggle Environment
------------------

Now that we have oooenv_ installed, we can toggle the virtual environment to be used by LibreOffice.

.. code-block:: powershell

    oooenv env -t

Output:

.. code-block:: text

    Saved cfg
    Saved cfg
    Set to UNO Environment

Now the environment is configured to be used by LibreOffice.

.. code-block:: powershell

    (.venv) PS D:\tmp\project> python --version
    Python 3.8.16

Run python in the virtual environment:

.. code-block:: powershell

    (.venv) PS D:\tmp\project> python

.. code-block:: python

    Python 3.8.16 (default, Apr 28 2023, 02:01:33) [MSC v.1929 64 bit (AMD64)] on win32
    Type "help", "copyright", "credits" or "license" for more information.
    >>> import uno

Toggle Environment:

When you want to switch back and forth to the original environment, run:

.. code-block:: powershell

    (.venv) PS D:\tmp\project> oooenv env -t
    Set to Original Environment

Install additional packages using poetry
----------------------------------------

Now that we have a virtual environment that can be used by LibreOffice, we can install additional packages using poetry_.

Make sure we are not in `UNO Environment`:

.. code-block:: powershell

    (.venv) PS D:\tmp\project> oooenv env -u
    NOT a UNO Environment

.. code-block:: powershell

    poetry add ooo-dev-tools

Output:

.. code-block:: text

    Using version ^0.11.6 for ooo-dev-tools

    Updating dependencies
    Resolving dependencies... (0.9s)

    Package operations: 6 installs, 0 updates, 0 removals

    • Installing types-uno-script (0.1.1)
    • Installing types-unopy (1.2.3)
    • Installing typing-extensions (4.6.3)
    • Installing lxml (4.9.2)
    • Installing ooouno (2.1.2)
    • Installing ooo-dev-tools (0.11.6)

    Writing lock file

Now we can see in our ``pyproject.toml`` file that the ``ooo-dev-tools`` (|odev|_) package has been added:

.. code-block:: toml
    :emphasize-lines: 3

    [tool.poetry.dependencies]
    python = "^3.8"
    ooo-dev-tools = "^0.11.6"

While we are in the original environment, we do not have access to LibreOffice and UNO.
So we will toggle again.

.. code-block:: powershell

    (.venv) PS D:\tmp\project> oooenv env -t
    Set to UNO Environment

Now we can take advantage of |odev|_.

.. code-block:: python

    >>> import uno
    >>> from ooodev.loader.lo import Lo
    >>> from ooodev.office.calc import Calc
    >>> from ooodev.utils.gui import GUI
    >>>
    >>> def say_hello(cell_name):
    >>>     sheet = Calc.get_active_sheet()
    ...     Calc.set_val(value="Hello World!", sheet=sheet, cell_name=cell_name)
    ...
    >>> _ = Lo.load_office(Lo.ConnectSocket())
    >>> doc = Calc.create_doc()
    >>> GUI.set_visible(visible=True, doc=doc)
    >>> say_hello("A1")
    >>> Lo.close_doc(doc)
    >>> Lo.close_office()

The result can be seen in :numref:`1cfcc990-9a1a-4117-964f-5df325dc437a`

.. cssclass:: screen_shot

    .. _1cfcc990-9a1a-4117-964f-5df325dc437a:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1cfcc990-9a1a-4117-964f-5df325dc437a
        :alt: Calc Hello World
        :figclass: align-center

        Calc Hello World

Recommended Python Packages
---------------------------

- ooo-dev-tools_ is a Python package that provides a framework to help with development of LibreOffice python projects. See |odev_docs|_.
- types-scriptforge_ is a Python package that provides type hints for the ScriptForge_ library.
- types-unopy_ is a Python package the has typings for the full LibreOffice API

.. note::

    Both ooo-dev-tools_ and types-scriptforge_ install the types-unopy_ package.

Related Links
-------------

- :ref:`guide_windows_manual_venv`
- :ref:`guide_lo_pip_windows_install`

.. _poetry: https://python-poetry.org/
.. _pyenv: https://github.com/pyenv-win/pyenv-win
.. _pip: https://pip.pypa.io/en/stable/
.. _oooenv: https://pypi.org/project/oooenv/
.. _ooo-dev-tools: https://pypi.org/project/ooo-dev-tools/
.. |odev_docs| replace:: OooDev Docs
.. _odev_docs: https://python-ooo-dev-tools.readthedocs.io/en/latest/index.html
.. _types-scriptforge: https://pypi.org/project/types-scriptforge/
.. _scriptforge: https://gitlab.com/LibreOfficiant/scriptforge
.. _types-unopy: https://pypi.org/project/types-unopy/