.. _guide_lo_pip_linux_install:

Linux - Install pip packages into LibreOffice
=============================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

Install pip_ into LibreOffice on Linux allows you to install python packages and use them in LibreOffice.

In most cases pip will already be installed on a linux system.

The process is essentially to install pip and then use it to install other python packages.

Steps
-----

Get LibreOffice Python Path
^^^^^^^^^^^^^^^^^^^^^^^^^^^

We need to know exactly where the system Python3 is located that LibreOffice is using.
Generally this will be ``/usr/bin/python3`` but it is best to check.
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

.. _guide_lo_pip_linux_install_test_pip:

Test Pip
^^^^^^^^

Make sure to use he same python executable that LibreOffice is using.

.. code-block:: bash

    $ /usr/bin/python3 -m pip --version
    pip 23.1.2 from /usr/local/lib/python3.10/dist-packages/pip (python 3.10)

If pip is not installed for some reason you can install it with the following command.

.. code:: bash

    curl -sSL https://bootstrap.pypa.io/get-pip.py | /usr/bin/python3 -


Install pip packages
^^^^^^^^^^^^^^^^^^^^

Using the command ``/usr/bin/python3 -m pip install ooo-dev-tools`` we can install the ooo-dev-tools_ package.
Notice that we are using the same python executable that LibreOffice is using.
Also notice that pip is installing the package into the user directory. ``Defaulting to user installation because normal site-packages is not writeable``.
This is because we are not running the command as root. This is usually the preferred method.
If for some reason you needed to install packages for all users you would need to run the command as root.
Don't do this unless you know what you are doing.

.. code-block:: bash
    :emphasize-lines: 2

    $ /usr/bin/python3 -m pip install ooo-dev-tools
    Defaulting to user installation because normal site-packages is not writeable
    Collecting ooo-dev-tools
    Using cached ooo_dev_tools-0.11.7-py3-none-any.whl (2.2 MB)
    Collecting lxml>=4.9.2 (from ooo-dev-tools)
    Using cached lxml-4.9.2-cp310-cp310-manylinux_2_17_x86_64.manylinux2014_x86_64.manylinux_2_24_x86_64.whl (7.1 MB)
    Collecting ooouno>=2.1.2 (from ooo-dev-tools)
    Using cached ooouno-2.1.2-py3-none-any.whl (9.8 MB)
    Collecting types-unopy>=1.2.3 (from ooouno>=2.1.2->ooo-dev-tools)
    Using cached types_unopy-1.2.3-py3-none-any.whl (5.2 MB)
    Collecting typing-extensions<5.0.0,>=4.6.2 (from ooouno>=2.1.2->ooo-dev-tools)
    Using cached typing_extensions-4.6.3-py3-none-any.whl (31 kB)
    Collecting types-uno-script>=0.1.1 (from types-unopy>=1.2.3->ooouno>=2.1.2->ooo-dev-tools)
    Using cached types_uno_script-0.1.1-py3-none-any.whl (9.3 kB)
    Installing collected packages: typing-extensions, types-uno-script, lxml, types-unopy, ooouno, ooo-dev-tools
    Successfully installed lxml-4.9.2 ooo-dev-tools-0.11.7 ooouno-2.1.2 types-uno-script-0.1.1 types-unopy-1.2.3 typing-extensions-4.6.3


Test installed package
----------------------

For a test we can write Hello World into a new Writer document.

Start LibreOffice Writer.
Using ``APSO`` console we can run the following script from within LibreOffice.
See: :ref:`guide_apso_installation`.

.. code-block:: python

    APSO python console [LibreOffice]
    3.8.16 (default, Apr 28 2023, 09:24:49) [MSC v.1929 32 bit (Intel)]
    Type "help", "copyright", "credits" or "license" for more information.
    >>> from ooodev.utils.lo import Lo
    >>> from ooodev.office.write import Write
    >>>
    >>> def say_hello():
    ...     cursor = Write.get_cursor(Write.active_doc)
    ...     Write.append_para(cursor=cursor, text="Hello World!")
    ... 
    >>> say_hello()
    >>> 


.. image:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/b370cae2-a6f6-41b7-9dfb-be6e4514bbf6
    :alt: Writer Hello World!
    :align: center

Related Links
-------------

- :ref:`guide_linux_manual_venv_snap`
- :ref:`guide_linux_poetry_venv`
- :ref:`guide_apso_installation`

.. _ooo-dev-tools: https://pypi.org/project/ooo-dev-tools/
.. _pip: https://pip.pypa.io/en/stable/