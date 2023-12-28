.. _help_draw_format_modify_indent_space_indent:

Draw Modify Indent
==================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.indent_space.Indent`, class is used to modify the values seen in :numref:`26662808-1f45-4dbd-885f-74fa3583edf4` of a style.

.. cssclass:: screen_shot

    .. _26662808-1f45-4dbd-885f-74fa3583edf4:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/26662808-1f45-4dbd-885f-74fa3583edf4
        :alt: Draw dialog Indent default
        :figclass: align-center
        :width: 450px

        Draw dialog Indent default


Setting the Indent
-------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 25, 26, 27, 28, 29, 30, 31, 32

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.format.draw.modify import FamilyGraphics, DrawStyleFamilyKind
        from ooodev.format.draw.modify.indent_space import Indent
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

                style = Indent(
                    before=4.5,
                    after=5.5,
                    first=5.2,
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

    .. _9f9d7f38-4b24-4196-b896-daf8ffa01a7c:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9f9d7f38-4b24-4196-b896-daf8ffa01a7c
        :alt: Draw dialog Indent style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Indent style changed

.. note::

    The ``auto`` argument of the ``Indent`` class constructor is suppose to set the styles ``ParaIsAutoFirstLineIndent`` property.
    The ``ParaIsAutoFirstLineIndent`` is suppose to be part of the ``com.sun.star.style.ParagraphProperties`` service;
    However, for some reason it is missing for Draw styles. Setting this ``auto`` argument will result
    in a print warning message in verbose mode. It is better to not set this argument.
    It is left in just in case it starts working in the future.

    There is a option in the Indent and Spacing dialog for ``Automatic`` as seen in :numref:`26662808-1f45-4dbd-885f-74fa3583edf4`.
    It seems to work, but it is not clear how it is implemented. It is not clear if it is a style property.


Getting indent from a style
---------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = Indent.from_style(
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
        - :ref:`help_draw_format_modify_indent_space_spacing`
        - :ref:`help_draw_format_modify_indent_space_line_spacing`
        - :py:class:`ooodev.format.draw.modify.indent_space.Indent`