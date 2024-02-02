.. _guide_ooodev_oxt_installation:

Install OooDev LibreOffice Extension
====================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 1

Overview
--------

.. image:: https://user-images.githubusercontent.com/4193389/260558706-6a975af5-0815-4d85-987a-6f8b3ff20609.png
    :width: 160px
    :alt: OooDev Tools
    :align: center

|odev| can also be installed as a LibreOffice extension.
When |odev| is installed as an extension it will be available to all of the LibreOffice Suite via python Macros.
The extension make is so you can use |odev| in any LibreOffice application without having to Pip install |odev| or `compile scripts <https://oooscript.readthedocs.io/en/latest/>`__ that use |odev|.

The OOO Development Tools extension is available on `GitHub <https://github.com/Amourspirit/libreoffice_ooodev_ext#readme>`__ and `LibreOffice Extensions <https://extensions.libreoffice.org/en/extensions/show/41700>`__.

.. note::

    This guide assumes you have already installed LibreOffice.

Download
--------

Download the extension and save it locally on your computer. The file will be named ``OooDev.oxt``.

Install
-------

From LibreOffice open the extension manager,  ``Tools -> Extension Manager``.

You should see a dialog similar to :numref:`4f6ff7e7-4ab9-4849-8c96-3cb0f4f527f1`.

.. cssclass:: screen_shot

    .. _4f6ff7e7-4ab9-4849-8c96-3cb0f4f527f1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4f6ff7e7-4ab9-4849-8c96-3cb0f4f527f1
        :alt: Extension Manager
        :figclass: align-center
        :width: 550px

        Extension Manager

Click Add and install the downloaded ``OooDev.oxt`` file.
The installed extension will be seen in the extension manager as seen in :numref:`88dd6878-7768-4e63-ae05-7b98b07714ae`.


.. cssclass:: screen_shot

    .. _88dd6878-7768-4e63-ae05-7b98b07714ae:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/88dd6878-7768-4e63-ae05-7b98b07714ae
        :alt: Extension Manager APSO Installed
        :figclass: align-center
        :width: 550px

        Extension Manager APSO Installed

You may be prompted to restart LibreOffice. Restart to use extension.

Example
-------

Using the APSO console (see :ref:`guide_apso_installation`) we can test the extension.

Write the following code in the APSO console.

.. tabs::

    .. code-tab:: python

        from ooodev.loader.lo import Lo
        from ooodev.office.write import Write
        from ooodev.macro.macro_loader import MacroLoader

        with MacroLoader():
            cursor = Write.get_cursor(Write.active_doc)
            Write.append_para(cursor=cursor, text="Hello World!")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

You should see results similar to :numref:`c63559f3-bbb6-45a5-8395-bbc4c0dd6079`.

.. cssclass:: screen_shot

    .. _c63559f3-bbb6-45a5-8395-bbc4c0dd6079:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/c63559f3-bbb6-45a5-8395-bbc4c0dd6079
        :alt: Alt
        :figclass: align-center

        :title

Related Links
-------------

- :ref:`guide_apso_installation`
