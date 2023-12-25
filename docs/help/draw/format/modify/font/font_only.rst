.. _help_draw_format_modify_font_font_only:

Draw Modify Font
================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.font.FontOnly`, class is used to modify the values seen in :numref:`15afed08-aae2-4361-b583-92b38b6810ea` of a style.

.. cssclass:: screen_shot

    .. _15afed08-aae2-4361-b583-92b38b6810ea:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/15afed08-aae2-4361-b583-92b38b6810ea
        :alt: Draw dialog Font default
        :figclass: align-center
        :width: 450px

        Draw dialog Font default


Setting the font
----------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 24, 25, 26, 27, 28, 29, 30, 31

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.format.draw.modify import FamilyGraphics, DrawStyleFamilyKind
        from ooodev.format.draw.modify.font import FontOnly, FontLang
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = DrawDoc(Draw.create_draw_doc())
                doc.set_visible()
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_75_PERCENT)

                slide = doc.get_slide()

                width = 100
                height = 50
                x = 10
                y = 10

                rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
                rect.set_string("Hello World!")
                font_style = FontOnly(
                    name="Arial",
                    size=20,
                    lang=FontLang().french_switzerland,
                    style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
                    style_family=DrawStyleFamilyKind.GRAPHICS,
                )
                doc.apply_styles(font_style)

                Lo.delay(1_000)
                doc.close_doc()
            return 0

        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following results in the Draw dialog.
Note: that the language is changed to French (Switzerland), this is optional via the :py:class:`~ooodev.format.inner.direct.write.char.font.font_only.FontLang` class.


.. cssclass:: screen_shot

    .. _14301e2c-faa0-43c8-b0e8-aa58daaafb08:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/14301e2c-faa0-43c8-b0e8-aa58daaafb08
        :alt: Draw dialog Font style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Font style changed


Getting font from a style
-------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = FontOnly.from_style(
            doc=doc.component,
            style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
            style_family=DrawStyleFamilyKind.GRAPHICS,
        )
        assert f_style.prop_style_name == str(FamilyGraphics.DEFAULT_DRAWING_STYLE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_draw_format_modify_font_font_effects`
        - :py:class:`ooodev.format.draw.modify.font.FontOnly`