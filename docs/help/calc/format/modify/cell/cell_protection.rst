.. _help_calc_format_modify_cell_protection:

Calc Modify Cell Protection
===========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.calc.modify.cell.cell_protection.CellProtection` class is used to programmatically set styles protection properties
as seen dialog  shown in :numref:`236623260-28e5e54e-6fd8-4085-8605-ca72d310005c`.

.. warning::
    Note that cell protection is not the same as sheet protection and cell protection is only enabled with sheet protection is enabled.
    See |lo_help_cell_protect|_ for more information.

.. cssclass:: screen_shot

    .. _236623260-28e5e54e-6fd8-4085-8605-ca72d310005c:

    .. figure:: https://user-images.githubusercontent.com/4193389/236623260-28e5e54e-6fd8-4085-8605-ca72d310005c.png
        :alt: Calc Format Cell dialog Cell Style Protection
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Cell Style Protection


Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 15, 16, 17, 18, 19, 20, 21, 22

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.modify.cell.cell_protection import CellProtection, StyleCellKind


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 100)

                style = CellProtection(
                    hide_all=False,
                    hide_formula=True,
                    protected=True,
                    hide_print=True,
                    style_name=StyleCellKind.DEFAULT,
                )
                style.apply(doc)

                style_obj = CellProtection.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
                assert style_obj.prop_style_name == str(StyleCellKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the style protection
----------------------------

.. tabs::

    .. code-tab:: python

        style = CellProtection(
            hide_all=False,
            hide_formula=True,
            protected=True,
            hide_print=True,
            style_name=StyleCellKind.DEFAULT,
        )
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236623690-864f6a2f-600a-4e66-a5d5-54f855c9efc1`.

.. cssclass:: screen_shot

    .. _236623690-864f6a2f-600a-4e66-a5d5-54f855c9efc1:

    .. figure:: https://user-images.githubusercontent.com/4193389/236623690-864f6a2f-600a-4e66-a5d5-54f855c9efc1.png
        :alt: Calc Format Cell dialog Cell Style Protection set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Cell Style Protection set

Getting cell protection from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = CellProtection.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
        assert style_obj.prop_style_name == str(StyleCellKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_direct_cell_cell_protection`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.calc.modify.cell.cell_protection.CellProtection`

.. |lo_help_cell_protect| replace:: LibreOffice Help - Cell Protection
.. _lo_help_cell_protect: https://help.libreoffice.org/latest/en-US/text/scalc/01/05020600.html