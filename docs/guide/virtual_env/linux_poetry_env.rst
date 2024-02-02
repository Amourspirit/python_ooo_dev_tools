.. _guide_linux_poetry_venv:

Linux - Manually Creating a Virtual Environment using Poetry
============================================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

This guide will walk you through the steps of manually creating a virtual using Poetry_ to work with LibreOffice.

.. note::

    This guide assumes you have already installed LibreOffice and poetry.

See also: :ref:`guide_configure_poetry`


On Linux it is simple to create a virtual environment using Poetry_ once you have the the path to the LibreOffice python.

See :ref:`guide_linux_manual_venv_get_python_path` for instructions on how to get the path to the LibreOffice python.

For this guide we will assume the path is ``/usr/bin/python3``.

Steps
------

Create Project Directory
^^^^^^^^^^^^^^^^^^^^^^^^

Create a Directory for your project.

In this case we will use a directory called ``myproject`` in the home directory.

.. code-block:: bash

    mkdir ~/myproject
    cd ~/myproject

Init using Poetry
^^^^^^^^^^^^^^^^^

Use poetry_ to initialize a project.

.. code-block:: powershell

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

Instruct Poetry to use specific python
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. code-block:: bash

    poetry env use /usr/bin/python3

Output:

.. code-block::

    Creating virtualenv project in /home/paul/myproject/.venv
    Using virtualenv: /home/paul/myproject/.venv

Activate the Virtual Environment
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. code-block:: bash

    source .venv/bin/activate

Install OOOENV
^^^^^^^^^^^^^^

The virtual environment has been created but it does not yet have access to ``uno.py`` and ``unohelper.py`` which are needed to use the LibreOffice API.

For this we will use the oooenv_ package.

Make sure you have activated the virtual environment.

oooenv_ is a Python package that allows you to auto configure a virtual environment to be used by LibreOffice.

Install oooenv_ in the virtual environment:

.. code-block:: powershell

    poetry add --group=dev oooenv

.. note::

    The ``--group=dev`` option is used because we only need oooenv_ for development purposes.
    This option instructs Poetry_ to only add oooenv_ to the ``dev-dependencies`` section of the ``pyproject.toml`` file.

Now that the package is installed we can use it to configure the virtual environment to use ``uno.py`` and ``unohelper.py``.

.. code-block:: bash

    oooenv cmd-link -a

Now the virtual environment is configured to use ``uno.py`` and ``unohelper.py``.

Install additional packages using poetry
----------------------------------------

Now that we have a virtual environment that can be used by LibreOffice,
we can install additional packages using poetry_ such as ooo-dev-tools_.

.. code-block:: bash

    poetry add ooo-dev-tools

Now we can take advantage of |odev|_.


Start python from our virtual environment.

.. code-block:: bash

    $ python

Run a simple test to make sure everything is working.

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

Related Links
-------------

- :ref:`guide_linux_manual_venv`
- :ref:`guide_apso_installation`

.. _oooenv: https://pypi.org/project/oooenv/
.. _ooo-dev-tools: https://pypi.org/project/ooo-dev-tools/
.. _poetry: https://python-poetry.org/