.. _guide_windows_manual_venv:

Windows - Manually Creating a Virtual Environment
=================================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

This guide will walk you through the steps of creating a virtual environment manually for LibreOffice in Windows.

While |win_pre_venv|_ are available, this guide is for those who want to create their own virtual environment.

Another option is to use the |py_path_ext|_ extension to add virtual environment paths to LibreOffice,
this would work with all LibreOffice versions after ``Version 7.0`` on all operating systems.

For the purpose of this guide we will assume your project directory is ``D:\tmp\manual``

.. note::

    This guide is for use with pip_ only. See Also :ref:`guide_windows_poetry_venv` for more information.

Steps
-----

Get Version of LibreOffice Python
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Get the version that the current installed LibreOffice is using.

.. tabs::

    .. code-tab:: powershell

        &"C:\Program Files\LibreOffice\program\python.exe" --version

    .. group-tab:: Cmd Shell

        .. code-block:: bat

            "C:\Program Files\LibreOffice\program\python.exe" --version

This will output something like: ``3.8.16``. We will need this value later.

Create Virtual Environment
^^^^^^^^^^^^^^^^^^^^^^^^^^

Create the virtual environment. It is important the ``--without-pip`` be included.

.. tabs::

    .. code-tab:: powershell

        cd D:\tmp\manual\
        python -m venv --without-pip .venv

    .. group-tab:: Cmd Shell

        .. code-block:: bat

            cd D:\tmp\manual\
            python -m venv --without-pip .venv

Now there will be a subdirectory ``D:\tmp\manual\.venv``.

Edit ``.venv\pyvenv.cfg``, use version found above ``3.8.16``.
So, if your found version is ``3.10.5`` then the ``version_info`` would read ``version_info = 3.10.5.final.0`` and so on.
The ``prompt`` line is completely optional and can be what you want.

.. code-block:: ini

    home = C:\Program Files\LibreOffice\program
    implementation = CPython
    version_info = 3.8.16.final.0
    virtualenv = 20.17.1
    include-system-site-packages = false
    base-prefix = C:\Program Files\LibreOffice\program\python-core-3.8.16
    base-exec-prefix = C:\Program Files\LibreOffice\program\python-core-3.8.16
    base-executable = C:\Program Files\LibreOffice\program\python.exe
    prompt = myproject_3.8.16

.. note::

    If ``include-system-site-packages = true`` then both ``C:\Users\guide\AppData\Roaming\Python\Python38\site-packages`` (if it exist) and ``C:\Program Files\LibreOffice\program\python-core-3.8.16`` will also be included on python's ``sys.path``.
    This is usually not needed.

Activate Virtual Environment
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: powershell

        .\.venv\Scripts\Activate.ps1

    .. group-tab:: Cmd Shell

        .. code-block:: bat

            .\.venv\Scripts\activate.bat

Install Pip
^^^^^^^^^^^

Install pip (virtual environment must be active)

.. tabs::

    .. code-tab:: powershell

        Invoke-WebRequest -Uri https://bootstrap.pypa.io/get-pip.py -UseBasicParsing).Content | python.exe -

    .. group-tab:: Cmd Shell

        .. code-block:: bat

            curl https://bootstrap.pypa.io/get-pip.py -o get-pip.py && type get-pip.py | python.exe -

Test by checking version:

.. tabs::

    .. code-tab:: powershell

        (myproject_3.8.16) PS D:\tmp\manual> python -m pip --version
        pip 23.1.2 from D:\tmp\manual\.venv\lib\site-packages\pip (python 3.8)

    .. group-tab:: Cmd Shell

        .. code-block:: bat

            (.venv) D:\tmp\manual>python -m pip --version
            pip 23.1.2 from D:\tmp\manual\.venv\lib\site-packages\pip (python 3.8)

Install extra python packages.

.. code-block:: powershell

    python -m pip install ooo-dev-tools

A test to see if it worked, see :ref:`guide_lo_pip_windows_install_testing_pkg`.

.. note::

    Note that it is import that pip be run with ``python -m pip`` to ensure the correct pip is being used.

Link other python packages
^^^^^^^^^^^^^^^^^^^^^^^^^^

Optionally link LibreOffice user python into virtual environment.

Deactivate current virtual environment.

.. code-block:: powershell

    deactivate

Find the user path (path that pip has been installed in):
``C:\Users\guide\AppData\Roaming\Python\Python38\site-packages`` where ``guide`` is your user name.

Create a file in ``\.venv\Lib\site-packages`` name ``libre_office_user_pkg.pth`` (name is not important as long as it ends with ``.pth``).
Open the file in a text editor and add the path to the user python packages.

The contents of the ``libre_office_user_pkg.pth`` , where ``guide`` is your username, are as follows:

.. code-block:: text

    C:\Users\guide\AppData\Roaming\Python\Python38\site-packages

Save and close the file.

Now when the virtual environment is activated the user python packages will be included on python's ``sys.path``.


Reactivate Virtual Environment

.. tabs::

    .. code-tab:: powershell

        .\.venv\Scripts\Activate.ps1

    .. group-tab:: Cmd Shell

        .. code-block:: bat

            .\.venv\Scripts\activate.bat    

Related Links
-------------

- :ref:`guide_lo_pip_windows_install`
- :ref:`guide_windows_poetry_venv`
- |win_pre_venv|_

.. |win_pre_venv| replace:: Pre-configured virtual environments for Windows
.. _win_pre_venv: https://github.com/Amourspirit/lo-support_file/tree/main/virtual_environments/windows
.. _pip: https://pypi.org/project/pip/

.. |py_path_ext| replace:: Include Python Path for LibreOffice
.. _py_path_ext: https://extensions.libreoffice.org/en/extensions/show/41996