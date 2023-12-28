.. _help_draw_format_modify_indent_space_spacing:

Draw Modify Spacing
===================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.indent_space.Spacing`, class is used to modify the values seen in :numref:`6c8354a2-a50b-4a43-917c-84f72e3b46c6` of a style.

.. cssclass:: screen_shot

    .. _6c8354a2-a50b-4a43-917c-84f72e3b46c6:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/6c8354a2-a50b-4a43-917c-84f72e3b46c6
        :alt: Draw dialog Spacing default
        :figclass: align-center
        :width: 450px

        Draw dialog Spacing default


Setting the Spacing
-------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 26, 27, 28, 29, 30, 31, 32

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.format.draw.modify import FamilyGraphics, DrawStyleFamilyKind
        from ooodev.format.draw.modify.indent_space import Spacing
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

                style = Spacing(
                    above=5,
                    below=5.5,
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

    .. _24609207-5660-4328-a95d-4718daad03b1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/24609207-5660-4328-a95d-4718daad03b1
        :alt: Draw dialog Spacing style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Spacing style changed

.. note::

    The ``style_no_space`` for the ``Spacing`` class constructor argument is suppose to set the styles ``ParaContextMargin`` property.
    The ``ParaContextMargin`` is suppose to be part of the ``com.sun.star.style.ParagraphProperties`` service;
    However, for some reason it is missing for Draw styles. Setting this ``style_no_space`` argument will result
    in a print warning message in verbose mode. It is better to not set this argument.
    It is left in just in case it starts working in the future.

    There is a option in the Indent and Spacing dialog ``Do not add space between paragraphs of the same style``
    as seen in :numref:`6c8354a2-a50b-4a43-917c-84f72e3b46c6`.
    It seems to work, but it is not clear how it is implemented. It is not clear if it is a style property.


Getting spacing from a style
----------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = Spacing.from_style(
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
        - :ref:`help_draw_format_modify_indent_space_line_spacing`
        - :py:class:`ooodev.format.draw.modify.indent_space.Spacing`