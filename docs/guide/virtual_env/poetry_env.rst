.. _guide_configure_poetry:

Configure a Poetry environment
==============================

Overview
--------

Poetry_ is a tool that facilitates creating a Python virtual environment based on the project dependencies. You can declare the libraries your project depends on, and Poetry will install and update them for you.

Project dependencies are recorded in the ``pyproject.toml`` file that specifies required packages, scripts, plugins, and URLs. See the pyproject_ reference for more information about its structure and format.

Install Poetry
--------------

1. Open Terminal (on macOS and Linux) or PowerShell (on Windows) and execute the following command:

.. tabs::

    .. group-tab:: macOs

        .. code-block:: bash

            curl -sSL https://install.python-poetry.org | python3 -
    
    .. group-tab:: Windows

        .. code-block:: powershell

            (Invoke-WebRequest -Uri https://install.python-poetry.org -UseBasicParsing).Content | py -
        
        .. note::

                If you have installed Python through the Microsoft Store, replace ``py`` with ``python`` in the command above.

    .. group-tab:: Linux

        .. code-block:: bash

            curl -sSL https://install.python-poetry.org | python3 -

2. On macOS and Windows, the installation script will suggest adding the folder with the poetry executable to the PATH variable. Do that by running the following command:

.. tabs::

    .. group-tab:: macOs

        .. code-block:: bash

            export PATH="/Users/tutorial/.local/bin:$PATH"
    
    .. group-tab:: Windows

        .. code-block:: powershell

            $Env:Path += ";C:\Users\tutorial\AppData\Roaming\Python\Scripts"; setx PATH "$Env:Path"

Don't forget to replace ``tutorial`` with your username!

3. To verify the installation, run the following command:

.. code-block:: bash

    poetry --version

You should see something like ``Poetry (version 1.5.0)``.

Post-Install
------------

By Default Poetry_ installs virtual environments in a subfolder of the user's home directory.
To change the default location to install in the projects current directory (``.venv``), set configuration option ``virtualenvs.in-project`` to ``true``:

.. code-block:: shell

    poetry config virtualenvs.in-project true

Also, it is possible to instruct Poetry_ to create its virtual environments in the project's root directory via a ``poetry.toml`` in the root directory of the project:

.. code-block:: toml

    [virtualenvs]
    in-project = true

.. _poetry: https://python-poetry.org/
.. _pyproject: https://python-poetry.org/docs/pyproject/