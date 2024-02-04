.. _help_calc_format_direct_cell_cell_protection:

Calc Direct Cell Protection
===========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`~.calc.Calc` class has several methods related to cell protection and sheet protection.

    - :py:meth:`~.calc.Calc.get_cell_protection`
    - :py:meth:`~.calc.Calc.is_cell_protected`
    - :py:meth:`~.calc.Calc.protect_sheet`
    - :py:meth:`~.calc.Calc.unprotect_sheet`
    - :py:meth:`~.calc.Calc.is_sheet_protected`

The :py:class:`ooodev.format.calc.direct.cell.cell_protection.CellProtection` class is used to set the cell protection properties.

The :py:class:`~ooodev.format.calc.direct.cell.cell_protection.CellProtection` class gives you the same options
as Calc's Cell Protection Dialog tab, but without the dialog. as seen in :numref:`236297706-04be0d71-adec-44a9-804b-7849fccca40b`.

.. warning::
    Note that cell protection is not the same as sheet protection and cell protection is only enabled with sheet protection is enabled.
    See |lo_help_cell_protect|_ for more information.

.. cssclass:: screen_shot

    .. _236297706-04be0d71-adec-44a9-804b-7849fccca40b:

    .. figure:: https://user-images.githubusercontent.com/4193389/236297706-04be0d71-adec-44a9-804b-7849fccca40b.png
        :alt: Calc Format Cell dialog Cell Protection
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Cell Protection

Apply cell protection to a cell
-------------------------------

Setup
^^^^^

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 16, 17

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.direct.cell.cell_protection import CellProtection

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                cell = Calc.get_cell(sheet=sheet, cell_name="A1")
                style = CellProtection(hide_all=False, hide_formula=True, protected=True, hide_print=True)
                Calc.set_val(value="Hello", cell=cell, styles=[style])

                f_style = CellProtection.from_obj(cell)
                assert f_style is not None

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the cell protection
^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        style = CellProtection(hide_all=False, hide_formula=True, protected=True, hide_print=True)
        Calc.set_val(value="Hello", cell=cell, styles=[style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236298445-d62faac5-62b8-4e2f-a669-bc8e1f94710b`.

.. cssclass:: screen_shot

    .. _236298445-d62faac5-62b8-4e2f-a669-bc8e1f94710b:

    .. figure:: https://user-images.githubusercontent.com/4193389/236298445-d62faac5-62b8-4e2f-a669-bc8e1f94710b.png
        :alt: Calc Format Cell dialog Cell Protection set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Cell Protection set

Getting cell protection from a cell
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = CellProtection.from_obj(cell)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply cell protection to a range
--------------------------------

Setup
^^^^^

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 19, 20

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.direct.cell.cell_protection import CellProtection

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                sheet = Calc.get_sheet()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                Calc.set_val(value="Hello", sheet=sheet, cell_name="A1")
                Calc.set_val(value="World", sheet=sheet, cell_name="B1")
                rng = Calc.get_cell_range(sheet=sheet, range_name="A1:B1")

                style = CellProtection(hide_all=False, hide_formula=True, protected=True, hide_print=True)
                style.apply(rng)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the range protection
^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        style = CellProtection(hide_all=False, hide_formula=True, protected=True, hide_print=True)
        style.apply(rng)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236298445-d62faac5-62b8-4e2f-a669-bc8e1f94710b`.

Getting cell protection from a range
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

It is not recommended to get and instance of :py:class:`~ooodev.format.calc.direct.cell.cell_protection.CellProtection` from a range.
This is because a range can have multiple cells with different cell protection settings and the ``CellProtection`` may not properly represent the range.

.. |lo_help_cell_protect| replace:: LibreOffice Help - Cell Protection
.. _lo_help_cell_protect: https://help.libreoffice.org/latest/en-US/text/scalc/01/05020600.html

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_modify_cell_protection`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:meth:`Calc.get_cell_range() <ooodev.office.calc.Calc.get_cell_range>`
        - :py:meth:`Calc.get_cell() <ooodev.office.calc.Calc.get_cell>`
        - :py:meth:`Calc.get_cell_protection() <ooodev.office.calc.Calc.get_cell_protection>`
        - :py:meth:`Calc.is_cell_protected() <ooodev.office.calc.Calc.is_cell_protected>`
        - :py:meth:`Calc.protect_sheet() <ooodev.office.calc.Calc.protect_sheet>`
        - :py:meth:`Calc.unprotect_sheet() <ooodev.office.calc.Calc.unprotect_sheet>`
        - :py:meth:`Calc.is_sheet_protected() <ooodev.office.calc.Calc.is_sheet_protected>`