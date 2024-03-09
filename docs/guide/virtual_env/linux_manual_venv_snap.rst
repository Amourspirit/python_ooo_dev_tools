.. _guide_linux_manual_venv_snap:

Linux - Manually Creating a Virtual Environment for LibreOffice Snap
====================================================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

Unlike Windows, The Snap version LibreOffice on Linux does not use embedded Python. Instead, it use the system Python3.
This means Snap LibreOffice on Linux has access to full python.

Using :ref:`guide_zaz_pip_installation` would be the recommended way to install pip and python packages in LibreOffice.

Another option is to use the |py_path_ext|_ extension to add virtual environment paths to LibreOffice,
this would work with all LibreOffice versions after ``Version 7.0`` on all operating systems.

Note that this guide for of an ``snap`` installed version of LibreOffice. It does not cover the Flatpak version or the APT version.
See: :ref:`guide_linux_manual_venv`, :ref:`guide_linux_poetry_venv`.

Note: at the time of writing this guide, the Snap version of LibreOffice is considerably slower starting up than the APT version when starting with a script.

Prerequisites
-------------

The ``libreoffice-script-provider-python`` apt package must be installed. This package allows scripts to connect to LibreOffice.
On Windows this in not needed because LibreOffice embeds Python. However, on Linux LibreOffice requires it even for Snaps.

.. code-block:: bash

    sudo apt install libreoffice-script-provider-python

Steps
-----

.. _guide_linux_manual_venv_snap_get_python_path:

Get LibreOffice Python Path
^^^^^^^^^^^^^^^^^^^^^^^^^^^

We need to know exactly where the system Python3 is located that LibreOffice Snap is using.
Generally this will be ``/usr/bin/python3`` but it is best to check.
Perhaps the easiest way is to use the APSO extension for LibreOffice.
See :ref:`guide_apso_installation` for more information.

Start LibreOffice Snap and open the APSO extension. In this case we are using Writer.

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

The Virtual Environment must be activated to use the installed packages.

.. code-block:: bash

    source .venv/bin/activate

First we will install ooo-dev-tools_.

.. code-block:: bash

    python -m pip install ooo-dev-tools

For a test we will write a short script and run it on the command line.
This simple script starts python, Loads Snap LibreOffice Calc, and writes ``Hello World!`` into the first cell.
Then a message box is displayed asking if you want to close the document.

This script does a few extra things to make the virtual environment work with a snap installed version of LibreOffice.

First it gets the path to the virtual environment site packages directory in the current virtual environment.
Internally the ``get_virtual_env_site_packages_path()`` function uses ``VIRTUAL_ENV`` environment variable to build up the virtual environment path.
If you are running a special case where the virtual environment is not activated, set the ``VIRTUAL_ENV`` environment variable to the virtual environment path.

This is an example of a custom Build System For Sublime Text that uses the virtual environment.
If you are not using Sublime Text, you can ignore this.

.. code-block:: json

    {
        "selector": "source.python",
        "working_dir": "$project_path",
        "env": {"PYTHONPATH":".", "VIRTUAL_ENV": "./.venv"},
        "path":"$project_path/.venv/bin:$PATH",
        "cmd": ["$project_path.venv/bin/python", "-u", "$file"],
        "file_regex": "^[ ]*File \"(...*?)\", line ([0-9]*)"
    }


The ``PYTHONPATH`` environment variable is set to include the virtual environment ``site-packages`` directory.
This value is read By Snap LibreOffice to include any Python packages that are installed in the virtual environment.

By default |odev| will not look for LibreOffice in the snap directory.
For this reason we need to set the ``soffice`` path to the snap directory.

.. code-block:: python

    Lo.ConnectSocket(soffice="/snap/bin/libreoffice", env_vars={"PYTHONPATH": py_pth})

.. note::

    |odev| Also has an Environment Variable that can be set to the Path of LibreOffice.
    This is ``ODEV_CONN_SOFFICE``. If this environment variable is set then the ``soffice`` is not needed;
    However, the ``soffice`` parameter will override the environment variable.

.. note::

    Alternatively a script can be run the APSO console as seen in :ref:`guide_lo_portable_pip_windows_install_test`.

.. warning::

    Snap LibreOffice does not seem to allow connections if it started with a pipe connection.
    For this reason use ``Lo.ConnectSocket()`` to connect to Snap LibreOffice as seen in the example below.

.. code-block:: python

    from __future__ import annotations
    import uno
    from pathlib import Path
    from ooodev.calc import CalcDoc
    from ooodev.utils.kind.zoom_kind import ZoomKind
    from ooodev.loader import Lo
    from ooodev.utils import paths
    from ooodev.dialog.msgbox import (
        MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
    )


    def main() -> int:
        py_pth = paths.get_virtual_env_site_packages_path()
        # uncomment to include current working directory in PYTHONPATH
        # py_pth += f":{Path.cwd()}"
        _ = Lo.load_office(
            Lo.ConnectSocket(soffice="/snap/bin/libreoffice", env_vars={"PYTHONPATH": py_pth})
        )
        try:
            doc = CalcDoc.create_doc(visible=True)
            Lo.delay(500)
            doc.zoom(ZoomKind.ZOOM_100_PERCENT)

            sheet = doc.sheets[0]
            sheet["A1"]value = "Hello World!"

            msg_result = doc.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                doc.close()
                Lo.close_office()
            else:
                print("Keeping document open")

        except Exception:
            Lo.close_office()
            raise
        return 0


    if __name__ == "__main__":
        SystemExit(main())

The resulting document should look like this:

.. image:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1cfcc990-9a1a-4117-964f-5df325dc437a
    :alt: LibreOffice Calc Hello World
    :align: center
    :class: screen_shot

The System path shows that the virtual environment site packages directory is included.

.. code-block:: python
    :emphasize-lines: 8

    APSO python console [LibreOffice]
    3.10.6 (main, Mar 10 2023, 10:55:28) [GCC 11.3.0]
    Type "help", "copyright", "credits" or "license" for more information.
    >>> import sys
    >>> from pprint import pprint
    >>> pprint(sys.path)
    ['/snap/libreoffice/275/lib/libreoffice/program',
    '/home/guide/myproject/.venv/lib/python3.10/site-packages',
    '/snap/libreoffice/275/gnome-platform/usr/lib/python3/dist-packages',
    '/usr/lib/python310.zip',
    '/usr/lib/python3.10',
    '/usr/lib/python3.10/lib-dynload',
    '/home/guide/snap/libreoffice/275/.local/lib/python3.10/site-packages',
    '/usr/lib/python3/dist-packages',
    '/home/guide/snap/libreoffice/275/.config/libreoffice/4/user/uno_packages/cache/uno_packages/lu46534i9c.tmp_/apso.oxt/python/pythonpath']
    >>> 

Related Links
-------------

- :ref:`guide_linux_poetry_venv`
- :ref:`guide_lo_pip_linux_install`
- :ref:`guide_apso_installation`

.. _oooenv: https://pypi.org/project/oooenv/
.. _ooo-dev-tools: https://pypi.org/project/ooo-dev-tools/

.. |py_path_ext| replace:: Include Python Path for LibreOffice
.. _py_path_ext: https://extensions.libreoffice.org/en/extensions/show/41996