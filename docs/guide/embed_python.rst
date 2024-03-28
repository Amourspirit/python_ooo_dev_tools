.. _guide_embed_python_macro:

Guide on embedding python macros in a LibreOffice Document
==========================================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Introduction
------------

LibreOffice does not have a build in editor for python. This entire project was created using VS Code.

There are many advantages to using python for macros.
But as with all good things there are often drawbacks.

This guide address the drawback of getting python scripts into LibreOffice Documents.

One major drawback for Python macros in LibreOffice is the ability to have a macro that spans multiple files without having to build and install a python package.
This guide will address this issue with the help of |oooscript|_.


This guide requires that |oooscript|_ be installed. The recommended way to do this is to have a virtual environment set up to work in.
The recommended way to create an environment to work in is to actually use an existing preconfigured environment such as |llp|_. 

See Also
^^^^^^^^

- `Pip & Virtual Environments <https://python-ooo-dev-tools.readthedocs.io/en/latest/guide/virtual_env/index.html>`__.
- |llp|_.
- `Debug Macros in Vs Code <https://github.com/Amourspirit/live-libreoffice-python/wiki/Debug-Macros-in-Vs-Code>`__.
- :ref:`linux_linking_paths`.

Simple Macro
------------

The |python_sample|_ contains a very basic example.

This example has no external dependencies and only a single script file to embed.

.. code-block:: python

    from __future__ import unicode_literals

    def doc_dialog():
        """Display a doc-based dialog"""
        model = XSCRIPTCONTEXT.getDocument()
        smgr = XSCRIPTCONTEXT.getComponentContext().ServiceManager
        dp = smgr.createInstanceWithArguments("com.sun.star.awt.DialogProvider", (model,))
        dlg = dp.createDialog("vnd.sun.star.script:Standard.Dialog1?location=document")
        dlg.execute()
        dlg.dispose()


    g_exportedScripts = (doc_dialog,)


The ``g_exportedScripts = (doc_dialog,)`` is an instruction for LibreOffice to let know that ``doc_dialog`` and only ``doc_dialog`` should be listed as a macro.
If you have more macro methods they can be appended; such as, ``g_exportedScripts = (doc_dialog, my_macro)``.

Make
^^^^

The ``make build`` can be run in a terminal window for  the ``ex/general/python_sample/`` folder.

.. code-block:: bash

    (ooouno-ex-py3.11) root ➜ .../libreoffice_ex/ex/general/python_sample (develop) 
    $ make build
    oooscript compile --embed \
        --config "/workspace/libreoffice_ex/ex/general/python_sample/config.json" \
        --embed-doc "/workspace/libreoffice_ex/ex/general/python_sample/data/sample.odt" \
        --build-dir "build/python_sample"

As you can see the ``make build`` just call ``oooscript`` and passes it a few parameters.

- ``compile`` instructs ``ooooscript`` to runs its ``compile`` command.
- ``--embed`` is to embed the script into a document.
- ``--config`` is the config file that contains extra instructions.
- ``--embed-doc`` is the document to embed the script into.
- ``--build-dir`` is where to save the output. The output folder will be created if it does not exist.

Config
^^^^^^

The ``config.json`` file is a json file that contains extra instructions for the ``oooscript`` command.

.. code-block:: json

    {
        "id": "oooscript",
        "name": "python_sample",
        "app": "WRITER",
        "version": "1.0.0",
        "args": {
            "src_file": "sample.py",
            "output_name": "python_sample",
            "single_script": true
        }
    }

The above configuration sets which file should be embedded ``"src_file": "sample.py"`` the output name and that this is a single script (standalone).

Grid Example
------------

The |grid_sample|_ is another example of a stand alone script.
This example requires that |ooo_dev_tools_ext|_ be installed to run as a stand alone macro.

By running the ``make build`` command the ``grid_ex.py`` file is embedded into the ``data/sales_data.ods`` file and saved to ``/build/sales_grid/grid_dialog.ods``.

Code
^^^^

The ``grid_ex.py`` file contains the following at the bottom of the file.

.. code-block:: python

    # ... other code

    def show_grid(*args) -> None:
        doc = CalcDoc.from_current_doc()
        grid_ex = GridEx(doc=doc)
        grid_ex.show()

    g_exportedScripts = (show_grid,)

The ``g_exportedScripts = (show_grid,)`` is an instruction for LibreOffice to let know that ``show_grid`` and only ``show_grid`` should be listed as a macro.

The ``show_grid()`` method just creates an instance of the class above and call the ``show()`` method to display the dialog.

Config
^^^^^^

The ``config.json`` file is a json file that contains extra instructions for the ``oooscript`` command.

.. code-block:: json

    {
        "id": "oooscript",
        "name": "grid_dialog",
        "app": "CALC",
        "version": "1.0.0",
        "args": {
            "src_file": "grid_ex.py",
            "output_name": "grid_dialog",
            "single_script": true
        }
    }

The above configuration sets which file should be embedded ``"src_file": "grid_ex.py"`` the output name and that this is a single script (standalone).

Multi-script Macro
------------------

The |tabs_sample|_ is good example of a multi-script macro.
This example requires that |ooo_dev_tools_ext|_ be installed to run as a stand alone macro.

The macro depends on the following files:

- ``script.py``
- ``tab_dialog.py``
- ``listbox.py``
- ``listbox_multi_select.py``
- ``listbox_drop_down.py``

The ``script.py`` imports ``tab_dialog`` which in turn imports the other modules.

Code
^^^^

.. code-block:: python

    from ooodev.calc import CalcDoc
    from tab_dialog import Tabs


    def show_tabs(*args) -> None:
        doc = CalcDoc.from_current_doc()
        tabs = Tabs(doc=doc)
        tabs.show()

Note that the code above does not start with any ``from __future__`` imports such as ``from __future__ import annotations``.
In the main macro script (``script.py``) for a multi-script macro the ``from __future__ is not supported``;
However it is fine for sub-modules to have a ``from __future__ import ...``

Config
^^^^^^

The ``config.json`` file is a json file that contains extra instructions for the ``oooscript`` command.

.. code-block:: json

    {
        "id": "oooscript",
        "name": "tabs_list_box",
        "app": "CALC",
        "version": "1.0.0",
        "args": {
            "src_file": "script.py",
            "output_name": "tabs_list_box",
            "single_script": false,
            "clean": false,
            "exclude_modules": [
                "ooodev\\.*",
                "com\\.*",
                "ooo\\.*"
            ]
        },
        "methods": [
            "show_tabs"
        ]
    }

By default |oooscript|_ will look for all modules that a multi-script file uses and embed them into a single script.
Because this scripts depend on |ooo_dev_tools|_  (OooDev) which is provided also provided by |ooo_dev_tools_ext|_ for LibreOffice,
then we do not want to include any ``OooDev`` packages.
The ``exclude_modules`` section of the ``args`` exclude module that are part of ``OooDev``.

Embed
^^^^^

To compile the scripts into a single script and embed the output in a Calc document run the `make `build``.

.. code-block:: bash

    (ooouno-ex-py3.11) root ➜ .../libreoffice_ex/ex/dialog/tabs_list_box (develop) 
    $ make build
    oooscript compile --embed \
        --config "/workspace/libreoffice_ex/ex/dialog/tabs_list_box/config.json" \
        --embed-doc "/workspace/libreoffice_ex/ex/dialog/tabs_list_box/data/sales_data.ods" \
        --build-dir "build/sales_data"

As you can see ``make build`` calls ``oooscript``.
The output can be found in the ``build`` folder in the root of the project which in this case is ``/workspace/libreoffice_ex/build/sales_grid``.

Output Code
^^^^^^^^^^^

A copy of the file that is embedded in the document is also outputted along side of the Calc Document.
Below is a screenshot of the output ``tabs_list_box.py`` file. The code is partially cut off due to its length.
You can see in the screenshot that the main entry point ``scritp.py`` has its contents included at the end of the file.
The contents of the ``tab_dialog.py``, ``listbox.py``, ``listbox_drop_down.py`` and ``listbox_multi_select.py``
are also included and wrapped in ``__scriptmerge_write_module()`` methods.

Screenshot of code output.

.. cssclass:: screen_shot

    .. _b45e4718-2ce8-4088-8231-8b696acf5c15:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/b45e4718-2ce8-4088-8231-8b696acf5c15
        :alt: tabs_list_box.py output
        :figclass: align-center
        :width: 600px

        Calc Cell


FAQ
---

How does Multi-script Work?
^^^^^^^^^^^^^^^^^^^^^^^^^^^

As seen in :numref:`b45e4718-2ce8-4088-8231-8b696acf5c15` all the required code is contained in the output python file.
This is the file that gets embedded in a LibreOffice Document.

When |oooscript|_ is running it looks for all the imports that are contained within your module and sub-modules.
If the module is not a system module then |oooscript|_ will include it in the output file.
Any patterns that match in the ``exclude_modules`` section of the ``config.json`` are omitted.

Be aware that packages such as `Pandas <https://pandas.pydata.org/>`__ and `Numpy <https://numpy.org/>`__ have binaries as part of their code.
Packages that have binaries not supported.
If you need Pandas see `Pandas for LibreOffice <https://extensions.libreoffice.org/en/extensions/show/41998>`__ extension.
If you need Numpy see `Python Numpy <https://extensions.libreoffice.org/en/extensions/show/41995>`__ extension.

When the macro gets called the module gets imported.
A module is only imported once in a session ( unless a reload is manually called ).
When the module is imported it will automatically create a temp folder and write all the embedded modules into the temp folder.
Once the modules are written into the temp folder the path is added to the python system.
After the script is done the temp folder is deleted.

If your script has many dependent files then this will make for a larger file.
This may make for a few second delay when running the first call to a macro in the module.
After the first call the module will already be in memory and there will be no delay.
If you have a big module then consider consider loading the module on startup when the document loads.

Can my sub-modules be in a sub-folder?
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Yes. The |sudoku|_ uses this approach. The sub-module is in a sub-folder named ``lib``. 
The sub-module is imported into the the ``script.py`` file via ``from lib import sudoku_calc``

Where can I get more help on oooscript?
----------------------------------------

See |oooscript|_ docs.

Can I embed the ooo-dev-tools package in a document?
----------------------------------------------------

Yes. The |ooo_dev_tools|_ package is can be embedded in a document like any other dependencies.

In the |tabs_sample|_ There is a ``make build_ooodev`` options. The build uses the ``config_ooodev.json`` file.

.. code-block:: json

    {
        "id": "oooscript",
        "name": "tabs_list_box",
        "app": "CALC",
        "version": "1.0.0",
        "args": {
            "src_file": "script.py",
            "output_name": "tabs_list_box",
            "single_script": false,
            "clean": true,
            "exclude_modules": [
                "sphinx\\.*"
            ]
        },
        "methods": [
            "show_tabs"
        ]
    }

Note that ``"ooodev\\.*"``, ``"com\\.*"``, and ``"ooo\\.*"`` are not part of the ``exclude_modules``.
This means that the ``ooo-dev-tools`` package will be included in the output file.

Note it is recommended that the ``clean`` option is set to ``true`` when including packages with a lot of doc strings (the ``ooo-dev-tools`` package has a lot of doc strings).
The ``clean`` option will remove doc strings from the output file which can reduce the size of the output file.


.. |oooscript| replace:: oooscript
.. _oooscript: https://oooscript.readthedocs.io/en/latest/

.. |llp| replace:: Live LibreOffice Python
.. _llp: https://github.com/Amourspirit/live-libreoffice-python

.. |python_sample| replace:: Python Sample
.. _python_sample: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/general/python_sample

.. |grid_sample| replace:: Grid Example
.. _grid_sample: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/dialog/grid

.. |tabs_sample| replace:: Tab and List Box Dialog
.. _tabs_sample: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/dialog/tabs_list_box

.. |sudoku| replace:: LibreOffice Calc Sudoku
.. _sudoku: https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/calc/sudoku

.. |ooo_dev_tools| replace:: OOO Development Tools
.. _ooo_dev_tools: https://python-ooo-dev-tools.readthedocs.io/en/latest/

.. |ooo_dev_tools_ext| replace:: OOO Development Tools Extension
.. _ooo_dev_tools-ext: https://extensions.libreoffice.org/en/extensions/show/41700