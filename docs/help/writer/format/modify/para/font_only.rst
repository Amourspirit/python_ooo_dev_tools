.. _help_writer_format_modify_para_font_only:

Write Modify Paragraph Font
===========================

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

        from ooodev.format.writer.modify.para.font import FontOnly, FontLang
        from ooodev.format.writer.modify.para.font import StyleParaKind
        from ooodev.write import Write, WriteDoc, ZoomKind
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = WriteDoc(Write.create_doc())
                doc.set_visible()
                Lo.delay(300)
                doc.zoom(ZoomKind.ENTIRE_PAGE)

                para_font_style = FontOnly(
                    name="Arial",
                    size=20,
                    lang=FontLang().french_switzerland,
                    style_name=StyleParaKind.STANDARD,
                )
                doc.apply_styles(font_style)

                Lo.delay(1_000)

                Lo.close_doc(doc)
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

        f_style = FontOnly.from_style(
            doc=doc.component, style_name=StyleParaKind.STANDARD
        )
        assert f_style.prop_style_name == str(StyleParaKind.STANDARD)

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
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.para.font.FontOnly`