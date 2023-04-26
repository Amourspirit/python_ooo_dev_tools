.. _help_writer_format_modify_para_font_only:

Write Modify Paragraph Font Effects
===================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.font.FontOnly`, class is used to set the paragraph style font :numref:`234660447-f39415b9-ddda-46cc-be7e-1b3231f8350f`.

.. cssclass:: screen_shot

    .. _234660447-f39415b9-ddda-46cc-be7e-1b3231f8350f:
    .. figure:: https://user-images.githubusercontent.com/4193389/234660447-f39415b9-ddda-46cc-be7e-1b3231f8350f.png
        :alt: Writer dialog Paragraph Font default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Font default


Setting the font
----------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 13, 14, 15, 16, 17, 18, 19

        from ooodev.format.writer.modify.para.font import FontOnly, FontLang, StyleParaKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                para_font_style = FontOnly(
                    name="Arial",
                    size=20,
                    lang=FontLang().french_switzerland,
                    style_name=StyleParaKind.STANDARD,
                )
                para_font_style.apply(doc)

                style_obj = FontOnly.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
                Lo.delay(1_000)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following results in the Writer dialog.
Note: that the language is changed to French (Switzerland), this is optional via the :py:class:`~ooodev.format.inner.direct.write.char.font.font_only.FontLang` class.


.. cssclass:: screen_shot

    .. _234714325-06e76d31-d19d-413e-8bab-8c854b08f00a:
    .. figure:: https://user-images.githubusercontent.com/4193389/234714325-06e76d31-d19d-413e-8bab-8c854b08f00a.png
        :alt: Writer dialog Paragraph Font style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Font style changed


Getting font from a style
-------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = FontOnly.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_char_font_only`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.font.FontOnly`