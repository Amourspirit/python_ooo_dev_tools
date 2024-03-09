.. _linux_linking_paths:

Linux Linking Paths for Development
===================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 1


When developing on Linux, you may need to link to libraries that are not part of LibreOffice.
This guide is one of several ways that develop environments can be set up to work with LibreOffice.
See also |live_office|_ for a Codespace that is ready to go.

This guide is for Ubuntu 22.04, but should work for other versions of Ubuntu and other Linux distributions.

This guide requires that the |include_p_path|_ extension is installed into LibreOffice.

Setting up Virtual Environment
------------------------------

First, create a virtual environment for your project.

The path for the project root directory is ``~/Documents/Projects/Python/LibreOffice/demo_env`` for this demo.

.. code-block:: bash

    python3 -m venv venv
    source venv/bin/activate

Install `ooo-dev-tools <https://pypi.org/project/ooo-dev-tools/>`__.
This will install the `ooo-dev-tools` package, which has many useful tools for developing with LibreOffice and includes full typing (auto-complete) support for the LibreOffice API.

.. code-block:: bash

    pip install ooo-dev-tools


Add a module to your project
----------------------------

Create a new folder in your project called ``my_mod``.
In this folder create a file called ``__init__.py``. This file can be empty.
Create a file called ``hello.py`` and add the following code.

.. code-block:: python

    from __future__ import annotations
    from ooodev.loader import Lo

    def main():
        doc = Lo.current_doc    
        doc.msgbox('Hello, world!')

    if __name__ == '__main__':
        main()

This is simple code that will display a message box with the text "Hello, world!".

Linking to LibreOffice
----------------------

Using |include_p_path|_ is the easiest way to link to LibreOffice.

Start LibreOffice and open the extension manager ``Tools -> Extension Manager...`` Select ``LibreOffice Python Path`` Extension. Click on ``Options``.

.. _25afb530-2304-413d-aa44-121e4c249b92:

.. figure:: https://github.com/Amourspirit/libreoffice-python-path-ext/assets/4193389/25afb530-2304-413d-aa44-121e4c249b92
    :alt: Extension Manager
    :align: center

    Extension Manager

Select the ``Python Paths`` option page.

.. _4a739a95-f131-42c2-bb0b-c1aa73260b0b:

.. figure:: https://github.com/Amourspirit/libreoffice-python-path-ext/assets/4193389/4a739a95-f131-42c2-bb0b-c1aa73260b0b
    :alt: Python Paths
    :align: center

    Python Paths

Choose ``Add Folder`` and navigate to the Location of the ``site-packages`` for the virtual environment that was set up previously.
Also add the root directory for your project, ``demo_env`` in this case.

.. _981a52b5-1835-49b5-b0e4-a6cd3559538e:

.. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/981a52b5-1835-49b5-b0e4-a6cd3559538e
    :alt: Add Folder
    :align: center
    :width: 600

    Add Folder

After the folders have been added, click ``OK`` to close the dialog. Restart LibreOffice to apply the changes.

.. note::

    The path for the project root directory is ``~/Documents/Projects/Python/LibreOffice/demo_env`` for this demo.
    The ``site-packages`` folder is located in the virtual environment that was created earlier. The path to the ``site-packages`` folder is ``venv/lib/python3.10/site-packages``.

Running the ``hello`` module.

Open the APSO console. See :ref:`guide_apso_installation`.

Import your module and run the ``main`` function.

.. _3dbbec7c-2c26-4cdd-a9b8-fd1fa1da9176:

.. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3dbbec7c-2c26-4cdd-a9b8-fd1fa1da9176
    :alt: APSO Console
    :align: center

    APSO Console

.. code-block:: python

    from my_mod import hello
    hello.main()

Add Macro Support
-----------------

It may be useful to add macro support to your project. This can be done by creating a symbolic link to the project ``macros`` folder in the LibreOffice ``Script/python`` folder.

Create a folder called ``macros`` in the root of your project. This folder will contain the macro files.
Like the example above we will write a simple macro that will display a message box with the text "Hello, world!".

Create a file called ``say_hello.py`` in the ``macros`` folder and add the following code.

.. code-block:: python

    from __future__ import annotations
    from ooodev.loader import Lo


    def say_hello(*args):
        doc = Lo.current_doc
        doc.msgbox("Hello, world!")

    g_exportedScripts = (say_hello,)

LibreOffice uses the ``~/.config/libreoffice/4/user/Scripts/python`` folder to store Python macros. Create a symbolic link to the project ``macros`` folder in the ``~/.config/libreoffice/4/user/Scripts/python`` folder.

Make sure that the ``python`` folder exists in the ``~/.config/libreoffice/4/user/Scripts`` folder. If it does not exist, create it.

Run the following command to create the symbolic link.

.. code-block:: bash

    ln -s ~/Documents/Projects/Python/LibreOffice/demo_env/macros ~/.config/libreoffice/4/user/Scripts/python/my_macro

Now start LibreOffice and run the Macro.

``Tools -> Macros -> Run Macro...``

.. _d499a88c-d232-4daa-b3c7-d728386e5983:

.. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d499a88c-d232-4daa-b3c7-d728386e5983
    :alt: Run Macro
    :align: center

    Run Macro

Conclusion
----------

This guide has shown how to link to LibreOffice on Linux for development purposes.
With a link to the virtual environment ``site-packages`` and the project root directory, it is possible to develop and test Python code that uses the LibreOffice API.
The addition of a symbolic link to the project ``macros`` folder in the LibreOffice ``Script/python`` folder allows for the development and testing of Python macros.
While this is not the only way to set up a development environment for LibreOffice, it is a simple and effective way to get started.
When ever possible |live_office|_ is recommended for development.

.. note::

    This guide is for development purposes only. It is not recommended to use this method for production.

.. |include_p_path| replace:: Include Python Path for LibreOffice
.. _include_p_path: https://extensions.libreoffice.org/en/extensions/show/41996

.. |live_office| replace:: Live LibreOffice Python
.. _live_office: https://github.com/Amourspirit/live-libreoffice-python
