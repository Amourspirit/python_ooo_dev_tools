.. _guide_zaz_pip_installation:

Install Zaz-Pip LibreOffice Extension
=====================================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 1

Overview
--------

Pip_ is a package-management system written in Python used to install and manage software packages.
The Python Software Foundation recommends using pip for installing Python applications and their dependencies during deployment.
Pip_ connects to an online repository of public packages called the |pypi|_.

Note that an extension is not required to install Python packages. However, the extension does make it easier to install packages from within LibreOffice.
See :ref:`guide_lo_pip_windows_install` for more information.

The Zaz-Pip LibreOffice extension is a LibreOffice extension that allows pip installation of Python packages from within LibreOffice.

.. note::

    This guide assumes you have already installed LibreOffice.

Installation
------------

The Plugin can be found on `GitHub <https://git.cuates.net/elmau/zaz-pip/releases>`__.

Download the ``oxt`` version as seen in :numref:`49663d95-dbc9-4dca-9fb6-ea1fa5806fce`

.. cssclass:: screen_shot

    .. _49663d95-dbc9-4dca-9fb6-ea1fa5806fce:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/49663d95-dbc9-4dca-9fb6-ea1fa5806fce
        :alt: Zaz-Pip Release Download
        :figclass: align-center

        Zaz-Pip Release Download


Open Extension Manager

From LibreOffice open the extension manager,  ``Tools -> Extension Manager``.

.. cssclass:: screen_shot

    .. _a9a58e36-2c76-4fca-8b7d-95fc39ff863c:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a9a58e36-2c76-4fca-8b7d-95fc39ff863c
        :alt: LibreOffice Extension Manager
        :figclass: align-center
        :width: 550px

        LibreOffice Extension Manager

Click Add and install the downloaded ``ZAZPip_v0.10.0.oxt`` file.

.. cssclass:: screen_shot

    .. _a5712741-2778-4c05-bb18-6bd44febdfd9:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a5712741-2778-4c05-bb18-6bd44febdfd9
        :alt: LibreOffice Extension Manager with extension added
        :figclass: align-center
        :width: 550px

        LibreOffice Extension Manager with extension added

You may be prompted to restart LibreOffice. Restart to use extension.

Related Links
-------------

- :ref:`guide_pip_via_zaz_pip`
- :ref:`guide_lo_pip_windows_install`


.. _pip: https://pip.pypa.io/en/stable/

.. |pypi| replace:: Python Package Index
.. _pypi: https://pypi.org/