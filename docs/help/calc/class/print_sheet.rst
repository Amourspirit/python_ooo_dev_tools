.. _help_calc_module_class_print_sheet:

Help Calc.print_sheet() method
==============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Sometimes it is desirable to print a range of a spreadsheet document programmatically.
The :py:meth:`Calc.print_sheet() <ooodev.office.calc.Calc.print_sheet>` method can be used to accomplish this task.

Print a range of a spreadsheet document
---------------------------------------

Pass  parameters directly
^^^^^^^^^^^^^^^^^^^^^^^^^

Parameters a passed directly to :py:meth:`~ooodev.office.calc.Calc.print_sheet` method.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.office.calc import Calc

        # ... other code

        # open document
        doc = Calc.open_doc(fnm="my_spreadsheet.ods")

        # print range C6:G33 of the first sheet,
        # optionally idx value can be used to print from sheet matching idx number.
        Calc.print_sheet(printer_name="Brother MFC-L2750DW series", range_name="C6:G33")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Get the parameters from the user defined properties
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Using :py:meth:`Info.get_user_defined_props <ooodev.utils.info.Info.get_user_defined_props>` method we can get the
user defined properties seen in :numref:`c115250f-947c-413a-b06e-39473e4a421e` of the document.

The document custom properties contains the sheet index value to print and the printer name.

Using the :py:meth:`Calc.find_used_range() <ooodev.office.calc.Calc.find_used_range>` method we can get the used range of of the sheet.

.. tabs::

    .. code-tab:: python

        import uno
        from com.sun.star.beans import XPropertySet
        from ooodev.utils.lo import Lo
        from ooodev.office.calc import Calc
        from ooodev.utils.info import Info

        # ... other code

        doc = Calc.open_doc(fnm="my_spreadsheet.ods")
        user_props = Info.get_user_defined_props(doc)

        # get properties as XPropertySet
        ps = Lo.qi(XPropertySet, user_props, True)
        idx = int(ps.getPropertyValue("PrintSheet")) -1 # -1 because index starts at 0
        printer_name = ps.getPropertyValue("PrinterName")
    
        sheet = Calc.get_sheet(doc=doc, idx=idx)

        # find the used range of the sheet
        used_range = Calc.find_used_range(sheet=sheet) # XCellRange

        # send the used range to the printer
        Calc.print_sheet(printer_name=printer_name, cell_range=used_range)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`ch20_finding_with_cursors`
        - :ref:`help_common_modules_info_get_user_defined_props`