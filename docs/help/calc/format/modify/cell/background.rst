.. _help_calc_format_modify_cell_background:

Calc Modify Cell Background Color
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.calc.modify.cell.background.Color` class is used to set the background color of a style.

Apply the background color to a style
-------------------------------------

Setup
^^^^^

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.calc.modify.cell.background import Color as StyleBgColor, StyleCellKind
        from ooodev.utils.color import StandardColor


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                style = StyleBgColor(
                    color=StandardColor.BLUE_LIGHT2, style_name=StyleCellKind.DEFAULT
                )
                style.apply(doc)

                style_obj = StyleBgColor.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
                assert style_obj.prop_style_name == str(StyleCellKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the color
^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        style = StyleBgColor(
            color=StandardColor.BLUE_LIGHT2, style_name=StyleCellKind.DEFAULT
        )
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236554316-75c956fc-50b4-4f0e-8768-e836b8cee377`.

.. cssclass:: screen_shot

    .. _236554316-75c956fc-50b4-4f0e-8768-e836b8cee377:

    .. figure:: https://user-images.githubusercontent.com/4193389/236554316-75c956fc-50b4-4f0e-8768-e836b8cee377.png
        :alt: Calc dialog style Borders modified
        :figclass: align-center

        Calc dialog style Borders modified

Getting the color from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        style_obj = StyleBgColor.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
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
        - :ref:`help_calc_format_direct_cell_background`
        - :py:class:`ooodev.format.calc.modify.cell.background.Color`