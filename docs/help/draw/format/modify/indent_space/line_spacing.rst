.. _help_draw_format_modify_indent_space_line_spacing:

Draw Modify Line Spacing
========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.indent_space.LineSpacing`, class is used to modify the values seen in :numref:`f415d725-5f5c-4454-b380-80db1a9cbe2e` of a style.

.. cssclass:: screen_shot

    .. _f415d725-5f5c-4454-b380-80db1a9cbe2e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f415d725-5f5c-4454-b380-80db1a9cbe2e
        :alt: Draw dialog Line Spacing default
        :figclass: align-center
        :width: 450px

        Draw dialog Line Spacing default


Setting the Line Spacing
------------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 25, 26, 27, 28, 29, 30, 31

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.format.draw.modify import FamilyGraphics, DrawStyleFamilyKind
        from ooodev.format.draw.modify.indent_space import LineSpacing, ModeKind
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

                style = LineSpacing(
                    mode=ModeKind.PROPORTIONAL,
                    value=87,
                    style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
                    style_family=DrawStyleFamilyKind.GRAPHICS,
                )
                doc.apply_styles(style)

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

    .. _be655f75-5671-4d22-b25b-2485c09b8137:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/be655f75-5671-4d22-b25b-2485c09b8137
        :alt: Draw dialog Spacing style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Spacing style changed

Getting line spacing from a style
---------------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = LineSpacing.from_style(
            doc=doc.component,
            style_name=style.prop_style_name,
            style_family=style.prop_style_family_name
        )
        assert f_style is not None
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
        - :ref:`help_draw_format_modify_indent_space_indent`
        - :ref:`help_draw_format_modify_indent_space_spacing`
        - :py:class:`ooodev.format.draw.modify.indent_space.LineSpacing`