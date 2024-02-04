.. _guide_pip_via_zaz_pip_linux:

Linux - Install Pip Package via Zaz-Pip
=========================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 1

Overview
--------

Although pip_ can be installed into Linux LibreOffice manually, see :ref:`guide_lo_pip_linux_install`.
It is possible to pip install in LibreOffice using an extension.
See :ref:`guide_zaz_pip_installation`.

The steps below are the same for LibreOffice Snap. See also :ref:`guide_lo_portable_pip_windows_install`.

.. note::

    This guide assumes you have already installed LibreOffice.

    Anywhere you see ``guide`` it needs to be replaced with your username.

Installing Pip
--------------

Pip should already be installed in your Linux distribution.
If it is not, then see :ref:`guide_lo_pip_linux_install_test_pip` of the :ref:`guide_lo_pip_linux_install` guide.

Add a Pip package
-----------------

.. cssclass:: screen_shot

    .. _9ad0faf7-94e6-4afd-ace1-89381e4078a7:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9ad0faf7-94e6-4afd-ace1-89381e4078a7
        :alt: Admin PIP
        :figclass: align-center

        Admin PIP

We will install ooo-dev-tools_ package as an example.

Click the ``Admin PIP`` button.
Type in ``ooo-dev-tools``.

.. note::

    If you have a dark theme like this example, then you may not see the characters when you type them in ( white text on light yellow background).
    If this is the case, no worries, you can just select the text to see what it typed.

Enter package Name

.. cssclass:: screen_shot

    .. _f3336b49-e2a7-46e4-91b8-bfe3847cec5c:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f3336b49-e2a7-46e4-91b8-bfe3847cec5c
        :alt: Enter Package Name
        :figclass: align-center

        Enter Package Name

Click Yes to the popup see in :numref:`4193389/84b10f5d-27f9-4c01-8cf5-ab0465c57b6d`.

.. cssclass:: screen_shot

    .. _4193389/84b10f5d-27f9-4c01-8cf5-ab0465c57b6d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/84b10f5d-27f9-4c01-8cf5-ab0465c57b6d
        :alt: Enter Package Name
        :figclass: align-center

        Enter Package Name

Installing in this case did take a bit of time.
Be patient and wait to see ``Successfully installed ...`` as seen in :numref:`f9f269b9-4bee-4e60-9228-ce2f09d6bdd4`.

.. cssclass:: screen_shot

    .. _f9f269b9-4bee-4e60-9228-ce2f09d6bdd4:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f9f269b9-4bee-4e60-9228-ce2f09d6bdd4
        :alt: Enter Package Name
        :figclass: align-center

        Enter Package Name

Zaz_Pip installs it packages into user directory something like ``/home/guide/.local/lib/python3.10/site-packages`` where ``guide`` is the username.


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
    >>> from ooodev.office.write import Write
    >>>
    >>> def say_hello():
    ...     cursor = Write.get_cursor(Write.active_doc)
    ...     Write.append_para(cursor=cursor, text="Hello World!")
    ... 
    >>> say_hello()
    >>>

After the line ``Write.append_para(cursor=cursor, text="Hello World!")`` is added and the ``Enter`` key has been pressed,
``Hello World!`` will show up in the Writer document as seen in :numref:`5e69905c-1142-415f-86af-604e72982914_2`.
Now we have working pip packages and can add any pip package we need using ``Zaz-Pip`` extension.

.. cssclass:: screen_shot

    .. _5e69905c-1142-415f-86af-604e72982914_2:

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