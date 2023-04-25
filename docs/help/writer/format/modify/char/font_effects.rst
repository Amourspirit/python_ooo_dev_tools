.. _help_writer_format_modify_char_font_effects:

Write Modify Character FontEffects Class
========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.char.font.FontEffects` class is used to modify the values seen in :numref:`234149112-0424e8b3-edf8-4ca8-9be4-d695dee658e2` of a character style.


Before Settings

.. cssclass:: screen_shot

    .. _234149112-0424e8b3-edf8-4ca8-9be4-d695dee658e2:
    .. figure:: https://user-images.githubusercontent.com/4193389/234149112-0424e8b3-edf8-4ca8-9be4-d695dee658e2.png
        :alt: Writer dialog Character font effects default
        :figclass: align-center
        :width: 450px

        Writer dialog Character font effects default


Setting the font effects
------------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 13, 14

        from ooodev.format.writer.modify.char.font import FontEffects, StyleCharKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                font_style = FontEffects(color=StandardColor.BLUE_LIGHT1, underline=FontLine(line=FontUnderlineEnum.DOUBLE))
                font_style.apply(doc)

                cursor = Write.get_cursor(doc)
                Write.append_para(cursor=cursor, text="Hello World!")

                style_obj = FontEffects.from_style(doc=doc, style_name=StyleCharKind.SOURCE_TEXT)
                assert style_obj.prop_style_name == str(StyleCharKind.SOURCE_TEXT)
                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

After applying the font effects.

.. cssclass:: screen_shot

    .. _234149201-4d2a2df8-2610-4268-a76b-757a2eab5569:
    .. figure:: https://user-images.githubusercontent.com/4193389/234149201-4d2a2df8-2610-4268-a76b-757a2eab5569.png
        :alt: Writer dialog character font style changed
        :figclass: align-center
        :width: 450px

        Writer dialog character font style changed


Getting the font effects from a style
-------------------------------------

We can get the font effects from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = FontEffects.from_style(doc=doc, style_name=StyleCharKind.SOURCE_TEXT)
        assert style_obj.prop_style_name == str(StyleCharKind.SOURCE_TEXT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None



.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_char_font_only`
        - :ref:`help_writer_format_direct_char_font_effects`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.char.font.FontEffects`