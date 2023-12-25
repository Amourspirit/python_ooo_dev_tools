.. _help_draw_format_modify_font_font_effects:

Draw Modify Font Effects
========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.font.FontEffects`, class is used to modify the values seen in :numref:`63605a84-5460-46c2-9c6e-c0c90c2f775d` of a style.

.. cssclass:: screen_shot

    .. _63605a84-5460-46c2-9c6e-c0c90c2f775d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/63605a84-5460-46c2-9c6e-c0c90c2f775d
        :alt: Draw dialog Font Effects default
        :figclass: align-center
        :width: 450px

        Draw dialog Font Effects default


Setting the Font Effects
------------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 26, 27, 28, 29, 30, 31, 32 ,33

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.format.draw.modify import FamilyGraphics, DrawStyleFamilyKind
        from ooodev.format.draw.modify.font import FontEffects, FontUnderlineEnum
        from ooodev.format.draw.modify.font import FontLine
        from ooodev.utils.color import CommonColor
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = DrawDoc(Draw.create_draw_doc())
                doc.set_visible()
                Lo.delay(700)
                doc.zoom(ZoomKind.ZOOM_75_PERCENT)

                slide = doc.get_slide()

                width = 100
                height = 50
                x = 10
                y = 10

                rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
                rect.set_string("Hello World!")
                font_style = FontEffects(
                    color=CommonColor.RED,
                    underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.BLUE),
                    shadowed=True,
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

.. cssclass:: screen_shot

    .. _9cfa05bc-665d-4eff-a001-c9e13d4f6b56:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9cfa05bc-665d-4eff-a001-c9e13d4f6b56
        :alt: Draw dialog Font Effects style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Font Effects style changed

Shape after applying style.

.. cssclass:: screen_shot

    .. _0563cbd9-5dfc-408e-ab4e-e35d39275144:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/0563cbd9-5dfc-408e-ab4e-e35d39275144
        :alt: Shape after Style applied
        :figclass: align-center

        Shape after Style applied


Getting font effects from a style
---------------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = FontEffects.from_style(
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
        - :ref:`help_draw_format_modify_font_font_only`
        - :py:class:`ooodev.format.draw.modify.font.FontEffects`