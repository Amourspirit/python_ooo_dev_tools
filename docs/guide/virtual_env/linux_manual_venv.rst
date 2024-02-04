.. _guide_linux_manual_venv:

Linux - Manually Creating a Virtual Environment
===============================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

Unlike Windows, LibreOffice on Linux done not use embedded Python. Instead, it use the system Python3.
This means LibreOffice on Linux has access to full python.

Using :ref:`guide_zaz_pip_installation` would be the recommended way to install pip and python packages in LibreOffice.

Another option is to use the |py_path_ext|_ extension to add virtual environment paths to LibreOffice,
this would work with all LibreOffice versions after ``Version 7.0`` on all operating systems.

Note that this guide for of an ``apt`` installed version of LibreOffice. It does not cover the Flatpak version or the Snap version.
See: :ref:`guide_linux_manual_venv_snap`.

Prerequisites
-------------

The ``libreoffice-script-provider-python`` apt package must be installed. This package allows scripts to connect to LibreOffice.
On Windows this in not needed because LibreOffice embeds Python. However, on Linux LibreOffice requires it Even for Snaps.

.. code-block:: bash

    sudo apt install libreoffice-script-provider-python

Steps
-----

.. _guide_linux_manual_venv_get_python_path:

Get LibreOffice Python Path
^^^^^^^^^^^^^^^^^^^^^^^^^^^

We need to know exactly where the system Python3 is located that LibreOffice is using.
Perhaps the easiest way is to use the APSO extension for LibreOffice.
See :ref:`guide_apso_installation` for more information.

Start LibreOffice and open the APSO extension. In this case we are using Writer.

``Tools -> Macros -> Organize python scripts``

.. image:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/5010d2cc-8610-4874-a719-4cf6827ad8dc
    :alt: APSO Extension
    :align: center

Start the Python Console

.. code-block::

    APSO python console [LibreOffice]
    3.10.6 (main, May 29 2023, 11:10:38) [GCC 11.3.0]
    Type "help", "copyright", "credits" or "license" for more information.
    >>> import sys
    >>> sys.executable
    '/usr/bin/python3'
    >>> 

import ``sys`` and then use ``sys.executable`` to get the path to the system Python3.
In this case the system Python3 is located at ``/usr/bin/python3``.
This is the path needed to create the virtual environment.

Close the Python Console and Writer.

Create Project Directory
^^^^^^^^^^^^^^^^^^^^^^^^

Create a Directory for your project.

In this case we will use a directory called ``myproject`` in the home directory.

.. code-block:: bash

    mkdir ~/myproject
    cd ~/myproject

Create Virtual Environment
^^^^^^^^^^^^^^^^^^^^^^^^^^

In your project folder run the following command.

.. code-block:: bash

    /usr/bin/python3 -m venv .venv

If you get an error about ``python3.10-venv`` not being installed, then install it.

.. code-block:: bash

    sudo apt install python3.10-venv

Activate Virtual Environment
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Activate the virtual environment.

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

    python -m pip install oooenv

Now that the package is installed we can use it to configure the virtual environment to use ``uno.py`` and ``unohelper.py``.

.. code-block:: bash

    oooenv cmd-link -a

Now the virtual environment is configured to use ``uno.py`` and ``unohelper.py``.

Test installed package
----------------------

First we will install ooo-dev-tools_.

.. code-block:: bash

    python -m pip install ooo-dev-tools

For a test we can write Hello World into a new Writer document.

With ooo-dev-tools_ installed we can now run LibreOffice python right on the command line and interact with LibreOffice.
Alternatively run a script in the APSO console as seen in :ref:`guide_lo_portable_pip_windows_install_test`.
This simple script starts python, Loads LibreOffice Writer, and writes ``Hello World!``.

.. code-block:: python

    Python 3.10.6 (main, May 29 2023, 11:10:38) [GCC 11.3.0] on linux
    Type "help", "copyright", "credits" or "license" for more information.
    >>> from ooodev.loader.lo import Lo
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

The resulting document should look like this:

.. image:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/b370cae2-a6f6-41b7-9dfb-be6e4514bbf6
    :alt: LibreOffice Writer Hello World
    :align: center
    :class: screen_shot

Related Links
-------------

- :ref:`guide_linux_poetry_venv`
- :ref:`guide_apso_installation`

.. _oooenv: https://pypi.org/project/oooenv/
.. _ooo-dev-tools: https://pypi.org/project/ooo-dev-tools/

.. |py_path_ext| replace:: Include Python Path for LibreOffice
.. _py_path_ext: https://extensions.libreoffice.org/en/extensions/show/41996