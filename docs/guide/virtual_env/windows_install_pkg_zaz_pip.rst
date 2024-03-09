.. _guide_pip_via_zaz_pip:

Windows - Install Pip Package via Zaz-Pip
=========================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 1

Overview
--------

Although pip_ can be installed into Windows LibreOffice manually, see :ref:`guide_lo_pip_windows_install`.
It is possible to pip install in LibreOffice using an extension.
See :ref:`guide_zaz_pip_installation`.

Another option is to use the |py_path_ext|_ extension to add virtual environment paths to LibreOffice,
this would work with all LibreOffice versions after ``Version 7.0`` on all operating systems.

The steps below are the same for portable LibreOffice. See also :ref:`guide_lo_portable_pip_windows_install`.

.. note::

    This guide assumes you have already installed LibreOffice.

    Anywhere you see ``<username>`` it needs to be replaced with your Windows username.

Install Pip
-----------

Open Pip manager ``Tools -> Add Ons -> Open Pip``.

If pip is not installed, you will see a dialog like :numref:`8ac32491-5669-4f76-9a5e-2a593160f006`

.. cssclass:: screen_shot

    .. _8ac32491-5669-4f76-9a5e-2a593160f006:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/8ac32491-5669-4f76-9a5e-2a593160f006
        :alt: Pip Not Installed
        :figclass: align-center

        Pip Not Installed

Click Install PIP and Yes.

.. cssclass:: screen_shot

    .. _540ef929-4757-4d77-927c-67e4974d8761:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/540ef929-4757-4d77-927c-67e4974d8761
        :alt: Install PIP
        :figclass: align-center

        Install PIP

You may see an warning messages about not being on path as seen in :numref:`54c5c501-6533-4829-bae8-5029efdaf65f`. These message can be ignored.
Pip actually is installed at ``C:\Users\<username>\AppData\Roaming\Python\Python38\site-packages``, where ``<username>`` will be your Windows username.

.. cssclass:: screen_shot

    .. _54c5c501-6533-4829-bae8-5029efdaf65f:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/54c5c501-6533-4829-bae8-5029efdaf65f
        :alt: PIP Installed
        :figclass: align-center

        PIP Installed

After Pip is first installed, you may need to restart LibreOffice for it to pick up the new path.

Add a Pip package
-----------------

We will install ooo-dev-tools_ package as an example.

Click the ``Admin PIP`` button.
Type in ``ooo-dev-tools``.

.. note::

    If you have a dark theme like this example, then you may not see the characters when you type them in ( white text on light yellow background).
    If this is the case, no worries, you can just select the text to see what it typed.

Click Yes to the popup see in :numref:`3ec8eca0-500a-4adb-bf60-a0468f62c791`.

.. cssclass:: screen_shot

    .. _3ec8eca0-500a-4adb-bf60-a0468f62c791:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3ec8eca0-500a-4adb-bf60-a0468f62c791
        :alt: Enter Package Name
        :figclass: align-center

        Enter Package Name

Installing in this case did take a bit of time.
Be patient and wait to see ``Successfully installed ...`` as seen in :numref:`79737007-7b49-4a00-96ce-86933672787c`.

.. cssclass:: screen_shot

    .. _79737007-7b49-4a00-96ce-86933672787c:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/79737007-7b49-4a00-96ce-86933672787c
        :alt: Enter Package Name
        :figclass: align-center

        Enter Package Name


Testing installed Pip package
-----------------------------

ooo-dev-tools_ has been installed and now we can use it to do a quick test.

Open LibreOffice Writer.
Open the APSO console. See also :ref:`guide_apso_installation`.

Add each line to the APSO console, one line at a time followed by the ``Enter`` key.

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

After the line ``Write.append_para(cursor=cursor, text="Hello World!")`` is added and the ``Enter`` key has been pressed,
``Hello World!`` will show up in the Writer document as seen in :numref:`5e69905c-1142-415f-86af-604e72982914`.
Now we have working pip packages and can add any pip package we need using ``Zaz-Pip`` extension.

.. cssclass:: screen_shot

    .. _5e69905c-1142-415f-86af-604e72982914:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/5e69905c-1142-415f-86af-604e72982914
        :alt: Hello World!
        :figclass: align-center
        :width: 550px

        Hello World!


Related Links
-------------

- :ref:`guide_zaz_pip_installation`
- :ref:`guide_apso_installation`
- :ref:`guide_lo_portable_pip_windows_install`

.. _ooo-dev-tools: https://pypi.org/project/ooo-dev-tools/
.. _pip: https://pip.pypa.io/en/stable/

.. |py_path_ext| replace:: Include Python Path for LibreOffice
.. _py_path_ext: https://extensions.libreoffice.org/en/extensions/show/41996