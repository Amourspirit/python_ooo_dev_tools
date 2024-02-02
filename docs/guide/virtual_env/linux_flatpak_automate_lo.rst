.. _guide_linux_flatpak_automate_libreoffice:

Linux - Automate LibreOffice FlatPak 
====================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

This guide will show you how to automate LibreOffice FlatPak on Linux.

The process is to start LibreOffice FlatPak, wait for it to be ready, and then
send commands to it via a bridge connection using API.

Prerequisites
-------------

The ``libreoffice-script-provider-python`` apt package must be installed.
This package allows scripts to connect to LibreOffice.

.. code-block:: bash

    sudo apt install libreoffice-script-provider-python


Set up virtual environment
--------------------------

|lo_flatpak|_ uses a python that is included with FlatPak.

Get Python version
^^^^^^^^^^^^^^^^^^

To start, we need to get the major and minor version of Python from LibreOffice.

The FlatPak version of LibreOffice comes with APSO extension already installed.
Start FlatPak LibreOffice and open the APSO extension. In this case we are using Writer.

``Tools -> Macros -> Organize python scripts``

.. image:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/5010d2cc-8610-4874-a719-4cf6827ad8dc
    :alt: LibreOffice Flatpak APSO Extension
    :align: center

Start the Python Console

.. code-block:: python

    APSO python console [LibreOffice]
    3.10.11 (main, Nov 10 2011, 15:00:00) [GCC 12.2.0]
    Type "help", "copyright", "credits" or "license" for more information.
    >>> 

We can see in the console output above that the python version is ``3.10.11`` in this case.
Now we know we need to create a virtual environment for python ``3.10``.

Most Likely your system python will match the Major and the Minor version of python.

Find matching  Python version on your system
""""""""""""""""""""""""""""""""""""""""""""

.. code-block:: bash

    $ /usr/bin/python3 --version
    Python 3.10.6

If you do not get a match for ``python3`` then also try with major and minor prefix and suffix.
You system may have the version installed.

.. code-block:: bash

    $ /usr/bin/python3.10 --version
    Python 3.10.6

If you do not have a python version matching major and minor version, ``3.10`` in this case,
then it is recommended to use a tool such as pyenv_ to install a version of python that matches.

Create a virtual Environment
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Create a virtual environment for you system.

Make a project Directory.

.. code-block:: bash

    $ mkdir ~/my-project
    cd ~/my-project

Create Virtual Environment using the matching python version found above.

.. code-block:: bash

    /usr/bin/python3.10 -m venv .venv

Activate Virtual Environment
^^^^^^^^^^^^^^^^^^^^^^^^^^^^
.. code-block:: bash

    source .venv/bin/activate

Now that the virtual environment is activated, we can install the required packages.

Install OOOENV
--------------

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

Automate with ooo-dev-tools
---------------------------

Requires ooo-dev-tools_ version ``0.11.8`` or greater.

Install ooo-dev-tools_

.. code-block:: bash

    python -m pip install ooo-dev-tools

Run an automation script
^^^^^^^^^^^^^^^^^^^^^^^^

We will create a simple script to automate LibreOffice FlatPak by the Name of ``hello_world.py``.

Connecting to LibreOffice FlatPak is a little different than connecting to LibreOffice installed on the system.

First of all starting FlatPak LibreOffice from a python script seems to need the ``--display`` argument set to the current display.
If it is not set we may get an error and the script may hang.

So in the script below we are getting the current display and setting it to the ``--display`` argument via ``os.getenv('DISPLAY')``.

Next we must instruct |odev|_ to use the FlatPak version of LibreOffice.
Normally |odev|_ will use the system installed version of LibreOffice by default; However,
we need to instruct |odev|_ to use the FlatPak version of LibreOffice.

This is done by setting the ``soffice`` argument to the ``flatpak run org.libreoffice.LibreOffice/x86_64/stable`` value.
The display setting are passed to the ``extended_args`` argument.
It also seem ``ConnectPipe`` does not work with FlatPak LibreOffice so we are using ``ConnectSocket``.
The same thing is true with Snap installed LibreOffice.
The rest is straight forward for |odev|_.

.. code-block:: python

    from __future__ import annotations
    import os
    import uno
    from ooodev.office.calc import Calc
    from ooodev.utils.gui import GUI
    from ooodev.utils.kind.zoom_kind import ZoomKind
    from ooodev.loader.lo import Lo
    from ooodev.dialog.msgbox import (
        MsgBox, MessageBoxType, MessageBoxButtonsEnum, MessageBoxResultsEnum
    )


    def main():
        display_str = f"--display {os.getenv('DISPLAY', ':0')}"
        _ = Lo.load_office(
            Lo.ConnectSocket(
                soffice="flatpak run org.libreoffice.LibreOffice/x86_64/stable",
                extended_args=[display_str],
            )
        )
        try:
            doc = Calc.create_doc()
            GUI.set_visible(True, doc)
            Lo.delay(500)
            Calc.zoom(doc, ZoomKind.ZOOM_100_PERCENT)

            sheet = Calc.get_sheet(doc, 0)
            Calc.set_val(value="Hello World!", sheet=sheet, cell_name="A1")

            msg_result = MsgBox.msgbox(
                "Do you wish to close document?",
                "All done",
                boxtype=MessageBoxType.QUERYBOX,
                buttons=MessageBoxButtonsEnum.BUTTONS_YES_NO,
            )
            if msg_result == MessageBoxResultsEnum.YES:
                Lo.close_doc(doc=doc, deliver_ownership=True)
                Lo.close_office()
            else:
                print("Keeping document open")

        except Exception:
            Lo.close_office()
            raise
        return


    if __name__ == "__main__":
        main()

.. note::

    The Code to connection to LibreOffice FlatPak is something like this:

    .. code-block:: python

        # is it important to use shell=True
        import os
        from subprocess import Popen
        Popen(f'flatpak run org.libreoffice.LibreOffice/x86_64/stable --invisible --norestore --nofirststartwizard --nologo --accept="socket,host=localhost,port=2002,tcpNoDelay=1;urp;" --display {os.getenv("DISPLAY")}', shell=True)

Once the script is created and save as ``hello_world.py`` we can run it

.. code-block:: bash

    python hello_world.py

.. image:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/250aa5f6-b1ad-48de-b98a-ae9ab68cbb28
    :alt: hello_world
    :align: center

Related Links
-------------

- :ref:`guide_linux_flatpak_lo_pip`
- :ref:`guide_linux_manual_venv_snap`
- :ref:`guide_linux_manual_venv`

.. |lo_flatpak| replace:: LibreOffice Flatpak
.. _lo_flatpak: https://flathub.org/apps/org.libreoffice.LibreOffice

.. _oooenv: https://pypi.org/project/oooenv/
.. _ooo-dev-tools: https://pypi.org/project/ooo-dev-tools/
.. _pyenv: https://github.com/pyenv/pyenv#readme