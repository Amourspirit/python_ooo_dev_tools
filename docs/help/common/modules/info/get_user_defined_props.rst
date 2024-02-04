.. _help_common_modules_info_get_user_defined_props:

Help Info.get_user_defined_props() method
=========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 1

Overview
--------

LibreOffice allows you to assign custom information fields to your document as seen in :numref:`c115250f-947c-413a-b06e-39473e4a421e`

.. cssclass:: screen_shot

    .. _c115250f-947c-413a-b06e-39473e4a421e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/c115250f-947c-413a-b06e-39473e4a421e
        :alt: Custom properties of my_spreadsheet
        :figclass: align-center

        Custom properties of my_spreadsheet

:py:meth:`Info.get_user_defined_props <ooodev.utils.info.Info.get_user_defined_props>` method is used to get the user defined properties of a document.

The :py:meth:`Info.get_user_defined_props <ooodev.utils.info.Info.get_user_defined_props>` method returns a XPropertyContainer_ interface.

This method is used to get the user defined properties of a document.

Example
-------

.. tabs::

    .. code-tab:: python

        import uno
        from com.sun.star.beans import XPropertySet
        from ooodev.loader.lo import Lo
        from ooodev.office.calc import Calc
        from ooodev.utils.info import Info

        # ... other code

        doc = Calc.open_doc(fnm="my_spreadsheet.ods")
        user_props = Info.get_user_defined_props(doc)
        # get properties as XPropertySet
        ps = Lo.qi(XPropertySet, user_props, True)
        assert int(ps.getPropertyValue("PrintSheet")) == 2
        assert ps.getPropertyValue("PrinterName") == "Brother MFC-L2750DW series"

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_calc_module_class_print_sheet`
        - :py:meth:`ooodev.utils.info.Info.get_user_defined_props`

.. _XPropertyContainer: https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1beans_1_1XPropertyContainer.html