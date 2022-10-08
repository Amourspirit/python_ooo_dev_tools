Dev Docs
========

Virtual Environment
-------------------

|odev| use a virtual environment for development purposes.



In order to run test it is essential that ``uno.py`` and ``unohelper.py`` can be imported.
These files are part of the LibreOffice installation.
The location of these files vary depending on OS and other factors.

Linux
^^^^^

Set up virtual environment if not existing.

.. code-block:: text

    $ python -m venv ./env

Activate virtual environment and install development requirements.

.. code-block:: text

    (env) $ pip install -r requirements.txt

On Linux what is required to communicate with LibreOffice API it a copy of, or link to ``uno.py`` and ``unohelper.py`` in the virtual environment.
``uno.py`` sets up the necessary code that makes importing from LibreOffice API possible.

|odev| has a command to accomplish this in the virtual environment on Linux.

.. code-block:: text

    (env) $ python -m main cmd-link --add

After virtual environment is set up and **activated**, running the above command on Linux will search in known paths for ``uno.py`` and ``unohelper.py``
and create links to files in the current virtual environment.
That's it! Now should be ready for development.

For other options try:

    .. code-block:: text

        (env) $ python -m main cmd-link -h



Windows
^^^^^^^

Windows is a little trickery. Creating a link to ``uno.py`` and importing it will not work as it does in Linux.
This is due to the how LibreOffice implements the python environment on Windows.

The way |odev| works on Windows is a slight hack to the virtual environment.

Start by using terminal to create a ``venv`` environment in the projects root folder

.. code-block:: text

    PS C:\python_ooo_dev_tools> python -m venv ./env

Get LibreOffice python version.

.. code-block:: text

    PS C:\python_ooo_dev_tools> "C:\Program Files\LibreOffice\program\python.exe"
    Python 3.8.10 (default, Mar 23 2022, 15:43:48) [MSC v.1928 64 bit (AMD64)] on win32
    Type "help", "copyright", "credits" or "license" for more information.
    >>>

Edit ``env/pyvenv.cfg``  file.

.. code-block:: text

    PS C:\python_ooo_dev_tools> notepad .\env\pyvenv.cfg

Original may look something like:


.. code-block:: text

    home = C:\ProgramData\Miniconda3
    include-system-site-packages = false
    version = 3.9.7

Change to: With the version that is the same as current LibreOffice Version

.. code-block:: text

    home = C:\Program Files\LibreOffice\program
    include-system-site-packages = false
    version = 3.8.10

Testing Virtual Environment
^^^^^^^^^^^^^^^^^^^^^^^^^^^

For a quick test of environment import ``uno`` If there is no import  error you should be good to go.

.. code-block:: text

    (env) PS C:\python_ooo_dev_tools> python
    Python 3.8.10 (default, Mar 23 2022, 15:43:48) [MSC v.1928 64 bit (AMD64)] on win32
    Type "help", "copyright", "credits" or "license" for more information.
    >>> import uno
    >>>


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

        git commit -n -m "rename pip-env.txt to requirements.txt"

Docs
----

Building Docs
^^^^^^^^^^^^^

With virtual environment activated, open a terminal window and ``cd ./docs``

.. code-block:: text
    :caption: Linux

    (env) $ make html

.. code-block:: text
    :caption: Windows

    (env) PS > .\make.bat html

Viewing docs
^^^^^^^^^^^^

|online_docs|_ are available.
Viewing local docs can be done by starting a local webserver.

|odev| has a script tho make this easier. In a separate terminal window run:

.. code-block:: text
    :caption: Linux

    (env) $ python cmds/run_http.py

.. code-block:: text
    :caption: Windows

    (env) PS > python .\cmds\run_http.py

This starts a web server on localhost. Docs can the be viewed at http://localhost:8000/docs/_build/html/index.html

Doc Style
^^^^^^^^^

Doc for project are in the ``./docs`` folder.
Docs follow a basic style guide. If you are making any changes to docs please consult the ``./docs/sytle_guide.txt`` for guidelines.

Doc Spelling
^^^^^^^^^^^^

Manual spell check
""""""""""""""""""

Document are spelled check before commit by default when `Hooks` are set up.

Manual spell check can be run in a ``./docs`` terminal Windows.

.. code-block:: text

    (env) $ sphinx-build -b spelling . _build


Spelling custom dictionaries
""""""""""""""""""""""""""""

Custom spelling dictionaries are found in ``./docs/internal/dict/`` directory.
Any custom dictionary in this directory starting with ``spelling_*`` is auto-loaded into spellcheck.

.. |online_docs| replace:: Online Docs
.. _online_docs: https://python-ooo-dev-tools.readthedocs.io/en/latest/
