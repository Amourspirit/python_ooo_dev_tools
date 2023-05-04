.. _help_writer_format_modify_char_font_only:

Write Modify Character FontOnly Class
=====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.char.font.FontOnly` class is used to modify the font values seen in :numref:`234142867-3b8878ea-3542-47e2-bf97-2b82c9300743` of a character style.


Before Settings

.. cssclass:: screen_shot

    .. _234142867-3b8878ea-3542-47e2-bf97-2b82c9300743:
    .. figure:: https://user-images.githubusercontent.com/4193389/234142867-3b8878ea-3542-47e2-bf97-2b82c9300743.png
        :alt: Writer dialog Character font default
        :figclass: align-center
        :width: 450px

        Writer dialog Character font default


Setting the font name and size
------------------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 13, 14

        from ooodev.format.writer.modify.char.font import FontOnly, StyleCharKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                font_style = FontOnly(name="Consolas", size=14, style_name=StyleCharKind.SOURCE_TEXT)
                font_style.apply(doc)

                style_obj = FontOnly.from_style(doc=doc, style_name=StyleCharKind.SOURCE_TEXT)
                assert style_obj.prop_style_name == str(StyleCharKind.SOURCE_TEXT)

                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

After appling the font name and size

.. cssclass:: screen_shot

    .. _234143029-420b616b-d028-4691-884f-e6cbbc08535d:
    .. figure:: https://user-images.githubusercontent.com/4193389/234143029-420b616b-d028-4691-884f-e6cbbc08535d.png
        :alt: Writer dialog character font style changed
        :figclass: align-center
        :width: 450px

        Writer dialog character font style changed


Getting the font from a style
-----------------------------

We can get the font name and size from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = FontOnly.from_style(doc=doc, style_name=StyleCharKind.SOURCE_TEXT)
        assert style_obj.prop_style_name == str(StyleCharKind.SOURCE_TEXT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_char_font_effects`
        - :ref:`help_writer_format_direct_char_font_only`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.char.font.FontOnly`