.. _help_calc_format_modify_cell_font_only:

Calc Modify Cell Font Only
==========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.calc.modify.cell.font.FontOnly` class is used to modify the font values seen in :numref:`236439302-a3041f03-1109-41b8-9f76-bdaab12051d7` of a cell style.


Before Settings

.. cssclass:: screen_shot

    .. _236439302-a3041f03-1109-41b8-9f76-bdaab12051d7:

    .. figure:: https://user-images.githubusercontent.com/4193389/236439302-a3041f03-1109-41b8-9f76-bdaab12051d7.png
        :alt: Calc dialog style font default
        :figclass: align-center
        :width: 450px

        Calc dialog style font default


Setting the font name and size
------------------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 14, 15

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.modify.cell.font import FontOnly, StyleCellKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                font_style = FontOnly(name="Consolas", size=14, style_name=StyleCellKind.DEFAULT)
                font_style.apply(doc)

                style_obj = FontOnly.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
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

    .. _236440677-55e0dec8-3c14-4d17-a75b-a325e6a9f9b6:

    .. figure:: https://user-images.githubusercontent.com/4193389/236440677-55e0dec8-3c14-4d17-a75b-a325e6a9f9b6.png
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

        style_obj = FontOnly.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
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
        - :ref:`help_calc_format_direct_cell_font_only`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.calc.modify.cell.font.FontOnly`