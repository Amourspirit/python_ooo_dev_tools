.. _help_writer_format_direct_char_font_only:

Write Direct Character FontOnly Class
=====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.char.font.FontOnly` class gives you the same options
as Writer's Font Dialog, but without the dialog. as seen in :numref:`ss_writer_dialog_char_font`.

.. cssclass:: screen_shot invert

    .. _ss_writer_dialog_char_font:
    .. figure:: https://user-images.githubusercontent.com/4193389/229371613-9a968d75-ca48-44cf-b227-e88d1266a8a8.png
        :alt: Writer dialog Character font
        :figclass: align-center
        :width: 450px

        Writer dialog: Character font

Setting the font name and size
------------------------------

.. tabs::

    .. code-tab:: python

        import sys
        from ooodev.format.writer.direct.char.font import FontOnly
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.PAGE_WIDTH)
                cursor = Write.get_cursor(doc)
                font_style = FontOnly(name="Liberation Serif", size=20)
                Write.append_para(
                    cursor=cursor, text="Hello World!", styles=[font_style]
                )
                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0

        if __name__ == "__main__":
            sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`ss_writer_hello_world_char_font_20pt`.

.. cssclass:: screen_shot

    .. _ss_writer_hello_world_char_font_20pt:
    .. figure:: https://user-images.githubusercontent.com/4193389/229377797-5419cf8c-96cc-4b47-94c2-4b5923da3553.png
        :alt: Writer dialog Character font 20pt
        :figclass: align-center

        Hello World! with font Liberation Serif 20pt

It can bee seen in :numref:`ss_writer_dialog_char_font_20pt` that the font size is 20pt.

.. cssclass:: screen_shot invert

    .. _ss_writer_dialog_char_font_20pt:
    .. figure:: https://user-images.githubusercontent.com/4193389/229377833-6bd6a752-35ea-4daa-9a3c-5d08b7dfc7fa.png
        :alt: Writer dialog Character font 20pt
        :figclass: align-center

        Writer dialog: Character font 20pt


Getting the font name and size from the document
------------------------------------------------

Continuing from the code example above, we can get the font name and size from the document.

A paragraph cursor object is used to select the first paragraph in the document.
The paragraph cursor is then used to get the font style.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 7

        # ... other code

        para_cursor = Write.get_paragraph_cursor(cursor)
        para_cursor.gotoPreviousParagraph(False)
        para_cursor.gotoEndOfParagraph(True)

        font_style = FontOnly.from_obj(para_cursor)

        assert font_style.prop_name == "Liberation Serif"
        assert font_style.prop_size.value == 20
        para_cursor.gotoEnd(False)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None



.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_char_font_effects`
        - :ref:`help_writer_format_direct_char_font`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.char.font.FontOnly`