.. _guide_apso_installation:

Install APSO LibreOffice Extension
==================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 1

Overview
--------

APSO will install a macro organizer dedicated to python scripts.
The APSO extension is found on `LibreOffice Extensions <https://extensions.libreoffice.org/en/extensions/show/apso-alternative-script-organizer-for-python>`__.

.. cssclass:: screen_shot

    .. _ee0919a1-0ea8-49fc-8df5-47d1290f8750:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ee0919a1-0ea8-49fc-8df5-47d1290f8750
        :alt: APSO Extension
        :figclass: align-center

        APSO Extension

.. note::

    This guide assumes you have already installed LibreOffice.

Download
--------

Download the extension. The download button is lower on the page as seen in :numref:`e8cd598e-8dd8-4f08-b5bf-103b6ce48ec7`.


.. cssclass:: screen_shot

    .. _e8cd598e-8dd8-4f08-b5bf-103b6ce48ec7:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/e8cd598e-8dd8-4f08-b5bf-103b6ce48ec7
        :alt: Download APSO Extension
        :figclass: align-center

        Download APSO Extension

Install
-------

From LibreOffice open the extension manager,  ``Tools -> Extension Manager``.

.. cssclass:: screen_shot

    .. _475e433c-8698-4409-9670-5da6d96113d6:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/475e433c-8698-4409-9670-5da6d96113d6
        :alt: Extension Manager
        :figclass: align-center
        :width: 550px

        Extension Manager

Click Add and install the downloaded ``apso.oxt`` file.
The installed extension will be seen in the extension manager as seen in :numref:`2cc62c5f-0b45-4d0c-b8fb-35a1b5147d9d`.


.. cssclass:: screen_shot

    .. _2cc62c5f-0b45-4d0c-b8fb-35a1b5147d9d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2cc62c5f-0b45-4d0c-b8fb-35a1b5147d9d
        :alt: Extension Manager APSO Installed
        :figclass: align-center
        :width: 550px

        Extension Manager APSO Installed

You may be prompted to restart LibreOffice. Restart to use extension.

Configure APSO
--------------

Open the APSO organizer, ``Tools -> Extension Manager``.

Select the APSO extension as seen in :numref:`2cc62c5f-0b45-4d0c-b8fb-35a1b5147d9d` and click the Options button.

Set the desired editor.

.. note::

    If you are using LibreOffice as a snap on Ubuntu then leave the editor blank. When you choose to edit Macro from APSO a popup will appear offering a choice of editors.


.. cssclass:: screen_shot

    .. _f5ab23ee-1bc9-4235-ba62-f75f91a2dff0:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f5ab23ee-1bc9-4235-ba62-f75f91a2dff0
        :alt: APSO Configuration
        :figclass: align-center
        :width: 550px

        APSO Configuration

.. _guide_apso_installation_start_apso:

Starting APSO
-------------

To start APSO, ``Tools -> Macros -> Organize Python Scripts``.

Click the Menu button and select ``Python Shell`` as seen in :numref:`90b900d9-008b-467b-90bf-13bdf70eda22`.

.. cssclass:: screen_shot

    .. _90b900d9-008b-467b-90bf-13bdf70eda22:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/90b900d9-008b-467b-90bf-13bdf70eda22
        :alt: APSO Configuration
        :figclass: align-center

        APSO Configuration

This will open the python shell as seen in :numref:`1f27ad3f-e736-4a35-a3da-00d654bdd38e`.

.. cssclass:: screen_shot

    .. _1f27ad3f-e736-4a35-a3da-00d654bdd38e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1f27ad3f-e736-4a35-a3da-00d654bdd38e
        :alt: APSO Console
        :figclass: align-center

        APSO Console

Related Links
-------------

- :ref:`guide_lo_pip_windows_install`
- :ref:`guide_lo_portable_pip_windows_install`