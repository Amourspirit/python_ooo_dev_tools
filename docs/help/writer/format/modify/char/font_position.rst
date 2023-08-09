.. _help_writer_format_modify_char_font_position:

Write Modify Character FontPosition Class
=========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.char.font.FontPosition` class is used to modify the values seen in :numref:`234252792-f4eaab74-7279-43f6-80d9-16f9407304c0` of a character style.


Before Settings

.. cssclass:: screen_shot

    .. _234252792-f4eaab74-7279-43f6-80d9-16f9407304c0:
    .. figure:: https://user-images.githubusercontent.com/4193389/234252792-f4eaab74-7279-43f6-80d9-16f9407304c0.png
        :alt: Writer dialog Character font position default
        :figclass: align-center
        :width: 450px

        Writer dialog Character font position default


Setting the font position
-------------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 13, 14, 15, 16, 17, 18, 19, 20, 21, 22

        from ooodev.format.writer.modify.char.font import FontPosition, FontScriptKind, StyleCharKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
           with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                font_style = FontPosition(
                    script_kind=FontScriptKind.SUBSCRIPT,
                    raise_lower=10,
                    rotation=270,
                    scale=85,
                    spacing=3,
                    pair=True,
                    style_name=StyleCharKind.EXAMPLE,
                )
                font_style.apply(doc)

                style_obj = FontPosition.from_style(doc=doc, style_name=StyleCharKind.EXAMPLE)
                assert style_obj.prop_style_name == str(StyleCharKind.EXAMPLE)
                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

After applying the font position.

.. cssclass:: screen_shot

    .. _234254594-a4e7d9f0-d358-4ba7-97d7-cf59c181035c:
    .. figure:: https://user-images.githubusercontent.com/4193389/234254594-a4e7d9f0-d358-4ba7-97d7-cf59c181035c.png
        :alt: Writer dialog character font position changed
        :figclass: align-center
        :width: 450px

        Writer dialog character font position changed


Getting the font position from a style
--------------------------------------

We can get the font position from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = FontPosition.from_style(doc=doc, style_name=StyleCharKind.EXAMPLE)
        assert style_obj.prop_style_name == str(StyleCharKind.EXAMPLE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_char_font_only`
        - :ref:`help_writer_format_direct_char_font_position`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.char.font.FontPosition`