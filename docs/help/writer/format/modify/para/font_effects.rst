.. _help_writer_format_modify_para_font_effects:

Write Modify Paragraph Font Effects
===================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.font.FontEffects`, class is used to set the paragraph style font :numref:`234707232-b0aca0fb-1e06-4d65-b7a5-3a3638cb950a`.

.. cssclass:: screen_shot

    .. _234707232-b0aca0fb-1e06-4d65-b7a5-3a3638cb950a:
    .. figure:: https://user-images.githubusercontent.com/4193389/234707232-b0aca0fb-1e06-4d65-b7a5-3a3638cb950a.png
        :alt: Writer dialog Paragraph Font Effects default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Font Effects default


Setting the font
----------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 15, 16, 17, 18, 19, 20, 21

        from ooodev.format.writer.modify.para.font import FontEffects, FontLine
        from ooodev.format.writer.modify.para.font import FontUnderlineEnum, StyleParaKind
        from ooodev.utils.color import CommonColor
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
           with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_font_effects_style = FontEffects(
                    color=CommonColor.RED,
                    underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.BLUE),
                    shadowed=True,
                    style_name=StyleParaKind.STANDARD,
                )
                para_font_effects_style.apply(doc)

                style_obj = FontEffects.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following results in the Writer dialog.

.. cssclass:: screen_shot

    .. _234708184-72e9ad07-1bc2-4200-8857-513743318a50:
    .. figure:: https://user-images.githubusercontent.com/4193389/234708184-72e9ad07-1bc2-4200-8857-513743318a50.png
        :alt: Writer dialog Paragraph Font Effects style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Font Effects style changed


Getting font effects from a style
---------------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = FontEffects.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_char_font_only`
        - :ref:`help_calc_format_direct_cell_font_effects`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.font.FontEffects`