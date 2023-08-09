.. _help_writer_format_modify_para_font_position:

Write Modify Paragraph Position
===============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.font.FontPosition`, class is used to set the paragraph style font :numref:`234710910-1b0015cf-a605-49d1-bb54-2b3c39714cad`.

.. cssclass:: screen_shot

    .. _234710910-1b0015cf-a605-49d1-bb54-2b3c39714cad:
    .. figure:: https://user-images.githubusercontent.com/4193389/234710910-1b0015cf-a605-49d1-bb54-2b3c39714cad.png
        :alt: Writer dialog Paragraph Position default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Position default


Setting the font effects
------------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 14, 15, 16, 17, 18, 19, 20, 21, 22, 23

        from ooodev.format.writer.modify.para.font import FontPosition, FontScriptKind
        from ooodev.format.writer.modify.para.font import CharSpacingKind, StyleParaKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_font_style = FontPosition(
                    script_kind=FontScriptKind.SUPERSCRIPT,
                    raise_lower=12,
                    rel_size=80,
                    rotation=90,
                    fit=True,
                    spacing=CharSpacingKind.LOOSE,
                    style_name=StyleParaKind.STANDARD,
                )
                para_font_style.apply(doc)

                style_obj = FontPosition.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
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

    .. _235550243-87c843e4-3747-4471-99e4-34962f47cddc:
    .. figure:: https://user-images.githubusercontent.com/4193389/235550243-87c843e4-3747-4471-99e4-34962f47cddc.png
        :alt: Writer dialog Paragraph Position style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Position style changed


Getting font position from a style
----------------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

       style_obj = FontPosition.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
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
        - :ref:`help_writer_format_direct_char_font_position`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.font.FontPosition`