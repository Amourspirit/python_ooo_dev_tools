.. _help_calc_format_direct_cell_cell_protection:

Calc Direct Cell Protection
===========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_protection()`` method is called to set the cell protection properties.

The ``style_protection()`` class gives you the same options as Calc's Cell Protection Dialog tab,
but without the dialog. as seen in :numref:`236297706-04be0d71-adec-44a9-804b-7849fccca40b_1`.

.. warning::
    Note that cell protection is not the same as sheet protection and cell protection is only enabled with sheet protection is enabled.
    See |lo_help_cell_protect|_ for more information.

.. cssclass:: screen_shot

    .. _236297706-04be0d71-adec-44a9-804b-7849fccca40b_1:

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

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(100)

                cell = sheet["A1"]
                cell.value = "Hello"
                cell.style_protection(
                    hide_all=False,
                    hide_formula=True,
                    protected=True,
                    hide_print=True,
                )

                f_style = cell.style_protection_get()
                assert f_style is not None

                Lo.delay(1_000)
                doc.close()
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

        cell = sheet["A1"]
        cell.value = "Hello"
        cell.style_protection(
            hide_all=False,
            hide_formula=True,
            protected=True,
            hide_print=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236298445-d62faac5-62b8-4e2f-a669-bc8e1f94710b_1`.

.. cssclass:: screen_shot

    .. _236298445-d62faac5-62b8-4e2f-a669-bc8e1f94710b_1:

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
        f_style = cell.style_protection_get()
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

        from __future__ import annotations
        import uno
        from ooodev.calc import CalcDoc
        from ooodev.loader import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = CalcDoc.create_doc(visible=True)
                sheet = doc.sheets[0]
                Lo.delay(500)
                doc.zoom_value(100)

                rng = sheet.rng("A1:B1")
                sheet.set_array(
                    values=[["Hello", "World"]], range_obj=rng
                )

                cell_rng = sheet.get_range(range_obj=rng)
                cell_rng.style_protection(
                    hide_all=False,
                    hide_formula=True,
                    protected=True,
                    hide_print=True,
                )

                Lo.delay(1_000)
                doc.close()
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

        # ... other code
        cell_rng = sheet.get_range(range_obj=rng)
        cell_rng.style_protection(
            hide_all=False,
            hide_formula=True,
            protected=True,
            hide_print=True,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236298445-d62faac5-62b8-4e2f-a669-bc8e1f94710b_1`.

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
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`