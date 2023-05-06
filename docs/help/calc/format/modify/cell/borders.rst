.. _help_calc_format_modify_cell_borders:

Calc Modify Cell Borders
========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

Calc has a dialog, as seen in :numref:`236538390-d7be8685-68d7-4998-b937-483992d387a4`, that sets cell borders for a style. Here we will only look at how to apply borders to a style.
For more details refer to :ref:`help_calc_format_direct_cell_borders` that has many examples of setting cell borders directly and the various options.

The :py:class:`ooodev.format.calc.modify.cell.borders.Borders` class is used to set the border values.

.. cssclass:: screen_shot

    .. _236538390-d7be8685-68d7-4998-b937-483992d387a4:

    .. figure:: https://user-images.githubusercontent.com/4193389/236538390-d7be8685-68d7-4998-b937-483992d387a4.png
        :alt: Calc dialog style Borders default
        :figclass: align-center
        :width: 450px

        Calc dialog style Borders default

Setup
-----

General function used to run these examples:

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.calc.modify.cell.borders import Borders, Padding, Side, StyleCellKind
        from ooodev.office.calc import Calc
        from ooodev.utils.color import CommonColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo


        def main() -> int:
            with Lo.Loader(Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(300)
                Calc.zoom_value(doc, 130)

                borders = Borders(
                    border_side=Side(color=CommonColor.BLUE),
                    padding=Padding(all=1.5),
                    style_name=StyleCellKind.DEFAULT,
                )
                borders.apply(doc)

                style_obj = Borders.from_style(doc, style_name=StyleCellKind.DEFAULT)
                assert style_obj is not None

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Setting Style Borders
---------------------

.. tabs::

    .. code-tab:: python

        borders = Borders(
            border_side=Side(color=CommonColor.BLUE),
            padding=Padding(all=1.5),
            style_name=StyleCellKind.DEFAULT,
        )
        borders.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Modifying Style Borders as shown in the code above results in the following:

.. cssclass:: screen_shot

    .. _236549985-27b4d84a-1f91-4156-aa9d-2d06034723da:

    .. figure:: https://user-images.githubusercontent.com/4193389/236549985-27b4d84a-1f91-4156-aa9d-2d06034723da.png
        :alt: Calc dialog style Borders modified
        :figclass: align-center

        Calc dialog style Borders modified

Getting Borders from style
---------------------------

.. tabs::

    .. code-tab:: python

        style_obj = Borders.from_style(doc, style_name=StyleCellKind.DEFAULT)
        assert style_obj is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_para_borders`
        - :ref:`help_writer_format_direct_table`
        - :ref:`help_calc_format_direct_cell_borders`
        - :py:class:`ooodev.format.calc.modify.cell.borders.Borders`
        - :py:class:`ooodev.format.calc.modify.cell.borders.Padding`
        - :py:class:`ooodev.format.calc.modify.cell.borders.Shadow`