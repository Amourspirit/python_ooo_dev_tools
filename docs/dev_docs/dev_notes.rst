.. _dev_doc:
Dev Docs
========

.. _dev_doc_virtulal_env:
Virtual Environment
-------------------

It is assumed `<https://github.com/Amourspirit/python_ooo_dev_tools>`__ has been cloned or unzipped to a folder.

`Poetry <https://python-poetry.org/>`_ is required to install this project in a development environment.

|odev| uses a virtual environment for development purposes.

.. _dev_doc_ve_linux:

Linux
^^^^^

.. code-block:: text

    $ python -m venv ./.venv

Activate virtual environment.

.. code-block:: shell

    source ./.venv/bin/activate

Install requirements using Poetry.

.. code-block:: shell

    (.venv) $ poetry install

In order to run test it is essential that ``uno.py`` and ``unohelper.py`` can be imported.
These files are part of the LibreOffice installation.
The location of these files vary depending on OS and other factors.


On Linux what is required to communicate with LibreOffice API it a copy of, or link to ``uno.py`` and ``unohelper.py`` in the virtual environment.
``uno.py`` sets up the necessary code that makes importing from LibreOffice API possible.

|odev| has a command to accomplish this in the virtual environment on Linux.

.. code-block:: text

    (.venv) $ python -m main cmd-link --add

After virtual environment is set up and **activated**, running the above command on Linux will search in known paths for ``uno.py`` and ``unohelper.py``
and create links to files in the current virtual environment.
That's it! Now should be ready for development.

For other options try:

    .. code-block:: text

        (.venv) $ python -m main cmd-link -h

.. _dev_doc_ve_windos:

Windows
^^^^^^^

Windows is a little trickery. Creating a link to ``uno.py`` and importing it will not work as it does in Linux.
This is due to the how LibreOffice implements the python environment on Windows.

The way |odev| works on Windows is a slight hack to the virtual environment.

Start by using terminal to create a ``.venv`` environment in the projects root folder


.. code-block:: text

    PS C:\python_ooo_dev_tools> python -m venv .\.venv

Get LibreOffice python version.

.. code-block:: text

    PS C:\python_ooo_dev_tools> "C:\Program Files\LibreOffice\program\python.exe" --version

Will return something like ``Python 3.8.10``

Edit ``pyvenv.cfg``  file.

.. code-block:: text

    PS C:\python_ooo_dev_tools> notepad .\.venv\pyvenv.cfg

Original may look something like:


.. code-block:: text

    home = C:\Users\User\AppData\Local\Programs\Python\Python310
    include-system-site-packages = false
    version = 3.10.1

Change ``home`` and ``version`` using the current LibreOffice Version:

.. code-block:: text

    home = C:\Program Files\LibreOffice\program
    include-system-site-packages = false
    version = 3.8.10

Activate Virtual environment.

.. code-block:: text

     PS C:\python_ooo_dev_tools> .\.venv\Scripts\Activate

Install requirements using Poetry.

.. code-block:: text

    (.venv) PS C:\python_ooo_dev_tools> poetry install

.. _dev_doc_ve_test:

Testing Virtual Environment
^^^^^^^^^^^^^^^^^^^^^^^^^^^

For a quick test of environment import ``uno`` If there is no import  error you should be good to go.

.. code-block:: text

    PS C:\python_ooo_dev_tools> .\.venv\scripts\activate
    (.venv) PS C:\python_ooo_dev_tools> python
    Python 3.8.10 (default, Mar 23 2022, 15:43:48) [MSC v.1928 64 bit (AMD64)] on win32
    Type "help", "copyright", "credits" or "license" for more information.
    >>> import uno
    >>>


.. _dev_doc_hooks:

Hooks
-----

|odev| uses git hooks to ensure document and test are building.

Pointing git to hooks is required for actions to run.

After virtual environment for |odev| is activated, run the following one time command.

.. code-block:: shell

    git config --local core.hooksPath .githooks/

After setting up hooks, commits and push runs their corresponding hooks before committing or pushing code to repo.

Sometimes it may be prudent to not run hooks, such as adding a text file for internal purposes.
In these cases run ``--no--verify`` flag of git.

Example git ``--no-verify`` command:

    .. code-block:: shell

        git commit -n -m "rename somefile.txt to myfile.txt"

.. _dev_doc_docs:

Docs
----

.. _dev_doc_docs_bulding:

Building Docs
^^^^^^^^^^^^^

With virtual environment activated, open a terminal window and ``cd ./docs``

.. code-block:: text
    :caption: Linux

    (.venv) $ make html

.. code-block:: text
    :caption: Windows

    PS C:\python_ooo_dev_tools\docs> make html

.. _dev_doc_docs_view:

Viewing docs
^^^^^^^^^^^^

|online_docs|_ are available.
Viewing local docs can be done by starting a local webserver.

|odev| has a script to make this easier. In a separate terminal window run:

.. code-block:: text
    :caption: Linux

    (.venv) $ python cmds/run_http.py

.. code-block:: text
    :caption: Windows

    PS C:\python_ooo_dev_tools> python .\cmds\run_http.py

This starts a web server on localhost. Docs can the be viewed at http://localhost:8000/docs/_build/html/index.html

.. _dev_doc_docs_style:

Doc Style
^^^^^^^^^

Doc for project are in the ``./docs`` folder.
Docs follow a basic style guide. If you are making any changes to docs please consult the ``./docs/sytle_guide.txt`` for guidelines.

.. _dev_doc_docs_spell:

Doc Spelling
^^^^^^^^^^^^

.. _dev_doc_docs_spell_check:

Manual spell check
""""""""""""""""""

Documents are spelled checked before commit by default when `Hooks` are set up.

Manual spell check can be run in a ``./docs`` terminal Windows.

.. code-block:: text

    (.venv) $ sphinx-build -b spelling . _build

.. _dev_doc_docs_spell_dict:

Spelling custom dictionaries
""""""""""""""""""""""""""""

Custom spelling dictionaries are found in ``./docs/internal/dict/`` directory.
Any custom dictionary in this directory starting with ``spelling_*`` is auto-loaded into spellcheck.

.. |online_docs| replace:: Online Docs
.. _online_docs: https://python-ooo-dev-tools.readthedocs.io/en/latest/
