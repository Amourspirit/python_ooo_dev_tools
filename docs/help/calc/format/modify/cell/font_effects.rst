.. _help_calc_format_modify_cell_font_effects:

Calc Modify Cell Font Effects
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.calc.modify.cell.font.FontEffects` class is used to modify the font values seen in :numref:`236445444-2061ec5a-e12b-4cfb-af4c-98fe8ee4393f` of a cell style.


Before Settings

.. cssclass:: screen_shot

    .. _236445444-2061ec5a-e12b-4cfb-af4c-98fe8ee4393f:

    .. figure:: https://user-images.githubusercontent.com/4193389/236445444-2061ec5a-e12b-4cfb-af4c-98fe8ee4393f.png
        :alt: Calc dialog style font default
        :figclass: align-center
        :width: 450px

        Calc dialog style font default


Setting the font name and size
------------------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 16, 17

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.calc.modify.cell.font import FontEffects, FontLine
        from ooodev.format.calc.modify.cell.font import StyleCellKind, FontUnderlineEnum
        from ooodev.utils.color import StandardColor

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                font_style = FontEffects(color=StandardColor.BLUE_LIGHT1, underline=FontLine(line=FontUnderlineEnum.DOUBLE))
                font_style.apply(doc)

                style_obj = FontEffects.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
                assert style_obj.prop_style_name == str(StyleCellKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

After applying the font name and size

.. cssclass:: screen_shot

    .. _236446623-2fd6396b-0053-49a5-9860-07708115ae9a:

    .. figure:: https://user-images.githubusercontent.com/4193389/236446623-2fd6396b-0053-49a5-9860-07708115ae9a.png
        :alt: Calc dialog style font default changed
        :figclass: align-center
        :width: 450px

        Calc dialog style font default changed


Getting the font from a style
-----------------------------

We can get the font name and size from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = FontEffects.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
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
        - :ref:`help_calc_format_direct_cell_font_effects`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.calc.modify.cell.font.FontEffects`