.. _guide_lo_portable_pip_windows_install:

Windows - Install pip into LibreOffice Portable
===============================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 1

Overview
--------
|lo_port|_ has its own python installed.


Install pip_ into |lo_port| on Windows allows you to install python packages and use them in LibreOffice.

Another option is to use the |py_path_ext|_ extension to add virtual environment paths to LibreOffice,
this would work with all LibreOffice versions after ``Version 7.0`` on all operating systems.

The process is essentially to install pip and then use it to install other python packages.

Install pip
-----------

|lo_port| already has python installed on Windows.
For this guide we will assume |lo_port| is installed at ``D:\Portables\PortableApps\LibreOfficePortable``

In a PowerShell terminal navigate to ``program`` directory for your |lo_port| installation.

.. code-block:: powershell

    cd "D:\Portables\PortableApps\LibreOfficePortable\App\libreoffice\program"

.. hint::

    In window file manager if you hold down the shift key while right clicking a folder then the popup menu should include ``Open PowerShell window here``.

In PowerShell run the following command to install pip

.. code-block:: powershell

    (Invoke-WebRequest -Uri https://bootstrap.pypa.io/get-pip.py -UseBasicParsing).Content | .\python.exe -

You may get a warning that the pip install location is not on that path. This warning can be ignored.

.. code-block:: powershell

    [D:\Portables\PortableApps\LibreOfficePortable\App\libreoffice\program\]
    >(Invoke-WebRequest -Uri https://bootstrap.pypa.io/get-pip.py -UseBasicParsing).Content | .\python.exe -
    Collecting pip
    Using cached pip-23.1.2-py3-none-any.whl (2.1 MB)
    Collecting setuptools
    Using cached setuptools-67.8.0-py3-none-any.whl (1.1 MB)
    Collecting wheel
    Using cached wheel-0.40.0-py3-none-any.whl (64 kB)
    Installing collected packages: wheel, setuptools, pip
    WARNING: The script wheel.exe is installed in 'D:\Portables\PortableApps\LibreOfficePortable\App\libreoffice\program\python-core-3.8.16\Scripts' which is not on PATH.
    Consider adding this directory to PATH or, if you prefer to suppress this warning, use --no-warn-script-location.
    WARNING: The scripts pip.exe, pip3.8.exe and pip3.exe are installed in 'D:\Portables\PortableApps\LibreOfficePortable\App\libreoffice\program\python-core-3.8.16\Scripts' which is not on PATH.
    Consider adding this directory to PATH or, if you prefer to suppress this warning, use --no-warn-script-location.
    Successfully installed pip-23.1.2 setuptools-67.8.0 wheel-0.40.0

Pip should have installed to ``D:\Portables\PortableApps\LibreOfficePortable\App\libreoffice\program\python-core-3.8.16\lib\site-packages``

Check pip version. A successful version check shows that ``pip`` is indeed on a path know to LibreOffice python.

.. code-block:: powershell

    >.\python.exe -m pip --version
    pip 23.1.2 from D:\Portables\PortableApps\LibreOfficePortable\App\libreoffice\program\python-core-3.8.16\lib\site-packages\pip (python 3.8)
    [D:\Portables\PortableApps\LibreOfficePortable\App\libreoffice\program\]

.. note::

    Pip may report that is is installed in a different location.
    Such as ``C:\Users\bigby\AppData\Roaming\Python\Python38\site-packages\pip``, where ``bigby`` is your username.
    This would most likely be because pip was installed for the Windows version of LibreOffice.
    Both LibreOffice and LibreOffice Portable share this path if they use the same python version.

    This also means any python packages installed in this location will be available to both LibreOffice and LibreOffice Portable.

Install a python package.
We will install ooo-dev-tools_ for testing. Note that is many take a few minutes to install.

.. code-block:: powershell

    >.\python.exe -m pip install ooo-dev-tools
    Collecting ooo-dev-tools
    Using cached ooo_dev_tools-0.11.6-py3-none-any.whl (2.2 MB)
    Collecting lxml>=4.9.2 (from ooo-dev-tools)
    Using cached lxml-4.9.2-cp38-cp38-win32.whl (3.5 MB)
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
    [D:\Portables\PortableApps\LibreOfficePortable\App\libreoffice\program\]

.. _guide_lo_portable_pip_windows_install_test:

Test installed package
----------------------

For a test we can write Hello World into a new Writer document.

Start LibreOffice Portable Writer.
Using ``APSO`` console we can run the following script from within LibreOffice.
See: :ref:`guide_apso_installation`.

.. code-block:: python

    APSO python console [LibreOffice]
    3.8.16 (default, Apr 28 2023, 09:24:49) [MSC v.1929 32 bit (Intel)]
    Type "help", "copyright", "credits" or "license" for more information.
    >>> from ooodev.loader.lo import Lo
    >>> from ooodev.write import WriteDoc
    >>>
    >>> def say_hello():
    ...     doc = WriteDoc.from_current_doc()
    ...     cursor = doc.get_cursor()
    ...     cursor.append_para(text="Hello World!")
    ... 
    >>> say_hello()
    >>> 

The resulting document should look like :numref:`b370cae2-a6f6-41b7-9dfb-be6e4514bbf6_2`


.. cssclass:: screen_shot

    .. _b370cae2-a6f6-41b7-9dfb-be6e4514bbf6_2:

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
- :ref:`guide_lo_pip_windows_install`
- |win_pre_venv|_

.. _ooo-dev-tools: https://pypi.org/project/ooo-dev-tools/
.. _pip: https://pip.pypa.io/en/stable/

.. |lo_port| replace:: LibreOffice Portable
.. _lo_port: https://portableapps.com/de/apps/office/libreoffice_portable

.. |win_pre_venv| replace:: Pre-configured virtual environments for Windows
.. _win_pre_venv: https://github.com/Amourspirit/lo-support_file/tree/main/virtual_environments/windows

.. |odev_docs| replace:: OooDev Docs
.. _odev_docs: https://python-ooo-dev-tools.readthedocs.io/en/latest/index.html
.. _types-scriptforge: https://pypi.org/project/types-scriptforge/
.. _scriptforge: https://gitlab.com/LibreOfficiant/scriptforge
.. _types-unopy: https://pypi.org/project/types-unopy/

.. |py_path_ext| replace:: Include Python Path for LibreOffice
.. _py_path_ext: https://extensions.libreoffice.org/en/extensions/show/41996