.. _guide_linux_flatpak_lo_pip:

Linux - Install pip packages into LibreOffice FlatPak
=====================================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

Add python packages to |lo_flatpak|_ installed version.

Using :ref:`guide_zaz_pip_installation` for Flatpak is not possible prior to ``Version 1.0.0`` because the Flatpak is sandboxed;
However, in ``Version 1.0.0`` this was corrected and it now works.
Using  Zaz-Pip LibreOffice extension is now the recommended way to install pip packages into LibreOffice Flatpak.
This guide is for those that which to do this manually.

It is also possible to install pip packages into the Flatpak LibreOffice by creating a virtual environment and 
linking that virtual environment to the Flatpak LibreOffice.

Another option is to use the |py_path_ext|_ extension to add virtual environment paths to LibreOffice,
this would work with all LibreOffice versions after ``Version 7.0`` on all operating systems.

When you see ``guide`` in path names below, replace with your user name.

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

We can confirm python by running python and checking the location.

Input command:

.. code-block:: bash

    $ python

Command Prompt:

.. code-block:: python

    Python 3.10.6 (main, May 29 2023, 11:10:38) [GCC 11.3.0] on linux
    Type "help", "copyright", "credits" or "license" for more information.
    >>> import sys
    >>> sys.executable
    '/home/guide/my-project/.venv/bin/python'
    >>> exit()

Other FlatPak Python Note
^^^^^^^^^^^^^^^^^^^^^^^^^

Running the FlatPak platform I found this to match. It is not all that important as we are really only interested in matching the major and minor version of python to create a virtual environment.

.. code-block:: text

    $ flatpak run org.freedesktop.Platform
    Similar installed refs found for ‘org.freedesktop.Platform’:

    1) runtime/org.freedesktop.Platform/x86_64/21.08 (system)
    2) runtime/org.freedesktop.Platform/x86_64/22.08 (system)

    Which do you want to use (0 to abort)? [0-2]: 2
    [📦 org.freedesktop.Platform ~]$ python --version
    Python 3.10.11

Including Virtual Environment in LibreOffice FlatPak
----------------------------------------------------

There are a couple of ways to do this.

On simple way is to use a app such as FlatSeal_ to set the ``PYTHONPATH`` environment variable.

Find virtual environment site-packages
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

With virtual environment activated in the terminal start python using the ``python`` command.
It is not necessary but we will use ``pprint`` to display ``sys.path`` to make it a little more readable.

Input command:

.. code-block:: bash

    $ python

Command Prompt:

.. code-block:: python

    Python 3.10.6 (main, May 29 2023, 11:10:38) [GCC 11.3.0] on linux
    Type "help", "copyright", "credits" or "license" for more information.
    >>> import sys
    >>> from pprint import pprint
    >>> pprint(sys.path)
    ['',
    '/usr/lib/python310.zip',
    '/usr/lib/python3.10',
    '/usr/lib/python3.10/lib-dynload',
    '/home/guide/my-project/.venv/lib/python3.10/site-packages']
    >>> exit()

We are interested in the path for ``site-packages``. Once we have that we are done with the terminal for now.

Add path using FlatSeal
"""""""""""""""""""""""

In FlatSeal_, a new ``PYTHONPATH`` environment variable needs to be added with the value we found for ``site-packages`` above.

.. code-block:: ini

    PYTHONPATH=my-project/.venv/lib/python3.10/site-packages

The ``/home/guide/`` part of the path can be left off.
If it is not included then it get automatically appended when LibreOffice runs

FlatSeal screenshot for LibreOffice settings:

.. image:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a0012ec1-fe56-47cb-8c8c-5c4f5e71dd0d
    :alt: FlatSeal Add PYTHONPATH
    :align: center

Notes on PYTHONPATH
^^^^^^^^^^^^^^^^^^^

If you need to add more then a single path use ``:`` to separate the paths.

In  some cases ``PYTHONPATH`` does not work correctly when a part of the path has a directory that start with ``.`` such as ``/home/guide/.local/lib/python3.10``.
One work around for this issue is to create a system link to a path that does not contain the ``.local`` part of the path.

.. code-block:: bash

    ln -s /home/guide/.local/lib/python3.10 /home/guide/local/lib/python3.10

Now ``PYTHONPATH`` can be set like so:

.. code-block:: ini

    PYTHONPATH=/home/guide/local/lib/python3.10

or

.. code-block:: ini

    PYTHONPATH=local/lib/python3.10

Now when new package are installed in the virtual environment via pip the are still available to LibreOffice because it is linked to the original path.

Some paths are black listed for usage in FlatPak's. See the `docs <https://docs.flatpak.org/en/latest/sandbox-permissions.html#filesystem-access>`__ for more information.

- These directories are blacklisted: ``/lib``, ``/lib32``, ``/lib64``, ``/bin``, ``/sbin``, ``/usr``, ``/boot``, ``/root``, ``/tmp``, ``/etc``, ``/app``, ``/run``, ``/proc``, ``/sys``, ``/dev``, ``/var``
- Exceptions from the blacklist: ``/run/media``
- These directories are mounted under ``/var/run/host``: ``/etc``, ``/usr``

See Also: `Python Environment Variables <https://docs.python.org/3.10/using/cmdline.html#environment-variables>`__.

Using a script to start LibreOffice FlatPak
-------------------------------------------

In some cases is it preferred to only temporally add current virtual environment to LibreOffice FlatPak.
This can be done with a python script.

There is a script on GitHub called |office_py|_ that does this for us.

Place the script in the root of your virtual environment.
Activate your virtual environment.

.. code-block:: bash

    source .venv/bin/activate

Now you can start LibreOffice FlatPak using the script with the Virtual Environment's path automatically added to the path.

.. code-block:: text

    usage: office.py [-h] [--invisible] [--nologo] [--minimized] [--norestore] [--headless] [--path-no-root] {writer,calc,draw,impress,math,base,global", "web,none}

    Office

    positional arguments:
    {writer,calc,draw,impress,math,base,global", "web,none}

    options:
    -h, --help            show this help message and exit
            --invisible   Starts in invisible mode. Neither the start-up logo nor the initial program window will be visible.
                            Application can be controlled, and documents and dialogs can be controlled and opened via the API. Using the
                            parameter, the process can only be ended using the taskmanager (Windows) or the kill command (UNIX-like systems).
                            It cannot be used in conjunction with --quickstart.
    --nologo              Disables the splash screen at program start.
    --minimized           Starts minimized. The splash screen is not displayed.
    --norestore           enables restart and file recovery after a system crash.
    --headless            Starts in "headless mode" which allows using the application without GUI.
                            This special mode can be used when the application is controlled by external clients via the API.
    --path-no-root        If set then the root path is not included in PYTHONPATH.

Starting a LibreOffice Flatpak app is rather simple.

This command will start Writer and include the virtual environments paths in the ``sys.path``.

.. code-block:: bash

    python office.py writer

We can see this in the APSO console after running the above command.

.. code-block:: python
    :emphasize-lines: 8, 9

    APSO python console [LibreOffice]
    3.10.11 (main, Nov 10 2011, 15:00:00) [GCC 12.2.0]
    Type "help", "copyright", "credits" or "license" for more information.
    >>> import sys
    >>> from pprint import pprint
    >>> pprint(sys.path)
    ['/app/libreoffice/program',
    '/home/guide/my-project/.venv/lib/python3.10/site-packages',
    '/home/guide/my-project',
    '/usr/lib/python310.zip',
    '/usr/lib/python3.10',
    '/usr/lib/python3.10/lib-dynload',
    '/app/lib/python3.10/site-packages',
    '/usr/lib/python3.10/site-packages',
    '/home/guide/.var/app/org.libreoffice.LibreOffice/config/libreoffice/4/user/uno_packages/cache/uno_packages/lu56bigt.tmp_/apso.oxt/python/pythonpath']
    >>> 

Testing a package
-----------------

For a test we will install ooo-dev-tools_ in our virtual environment.

Command with virtual environment active.

.. code-block:: bash

    python -m pip install ooo-dev-tools

Output:

.. code-block:: text

    Collecting ooo-dev-tools
    Using cached ooo_dev_tools-0.11.8-py3-none-any.whl (2.2 MB)
    Collecting ooouno>=2.1.2
    Using cached ooouno-2.1.2-py3-none-any.whl (9.8 MB)
    Collecting lxml>=4.9.2
    Using cached lxml-4.9.2-cp310-cp310-manylinux_2_17_x86_64.manylinux2014_x86_64.manylinux_2_24_x86_64.whl (7.1 MB)
    Collecting typing-extensions<5.0.0,>=4.6.2
    Using cached typing_extensions-4.6.3-py3-none-any.whl (31 kB)
    Collecting types-unopy>=1.2.3
    Using cached types_unopy-1.2.3-py3-none-any.whl (5.2 MB)
    Collecting types-uno-script>=0.1.1
    Using cached types_uno_script-0.1.1-py3-none-any.whl (9.3 kB)
    Installing collected packages: typing-extensions, types-uno-script, lxml, types-unopy, ooouno, ooo-dev-tools
    Successfully installed lxml-4.9.2 ooo-dev-tools-0.11.8 ooouno-2.1.2 types-uno-script-0.1.1 types-unopy-1.2.3 typing-extensions-4.6.3

Start Office
^^^^^^^^^^^^

If your path is already include via FlatSeal then you can just start LibreOffice Writer Normally.
If you are using the script method then run ``python office.py writer``

| When Writer loads, open the APSO Console.
| ``Tools -> Macros -> Organize python scripts``

.. image:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/5010d2cc-8610-4874-a719-4cf6827ad8dc
    :alt: LibreOffice Flatpak APSO Extension
    :align: center

Run the follow code

.. code-block:: python

    APSO python console [LibreOffice]
    3.10.11 (main, Nov 10 2011, 15:00:00) [GCC 12.2.0]
    Type "help", "copyright", "credits" or "license" for more information.
    >>> from ooodev.write import WriteDoc
    >>> def say_hello():
    ...     doc = WriteDoc.from_current_doc()
    ...     cursor = doc.get_cursor()
    ...     cursor.append_para(text="Hello World!")
    ...
    >>> say_hello()
    >>>

When ``say_hello()`` is called ``Hello World!`` is automatically written into the document.

.. image:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/8e3f1fcc-1a19-4189-b228-4b94d106b426
    :alt: LibreOffice Flatpak APSO Say Hello
    :align: center

Related Links
-------------

- :ref:`guide_linux_flatpak_automate_libreoffice`
- :ref:`guide_linux_manual_venv_snap`
- :ref:`guide_linux_manual_venv`



.. _pyenv: https://github.com/pyenv/pyenv#readme
.. _flatseal: https://flathub.org/apps/com.github.tchx84.Flatseal
.. _ooo-dev-tools: https://pypi.org/project/ooo-dev-tools/

.. |office_py| replace:: office.py
.. _office_py: https://gist.github.com/Amourspirit/1540a52f21c020a8190b468a3e9efc16

.. |lo_flatpak| replace:: LibreOffice Flatpak
.. _lo_flatpak: https://flathub.org/apps/org.libreoffice.LibreOffice

.. |py_path_ext| replace:: Include Python Path for LibreOffice
.. _py_path_ext: https://extensions.libreoffice.org/en/extensions/show/41996