.. _help_draw_format_modify_area_color:

Write Modify Draw Area Color
============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.area.Color` class is used to modify the values seen in :numref:`f335230c-b84f-4beb-962d-59c8db3561e0` of a style.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 27, 28, 29, 30, 31, 32

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.format.draw.modify.area.color import Color
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.modify.area import Color as FillColor
        from ooodev.format.draw.modify import FamilyGraphics, DrawStyleFamilyKind
        from ooodev.utils.color import StandardColor


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
                style_modify = FillColor(
                    color=StandardColor.LIME_LIGHT2,
                    style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
                    style_family=DrawStyleFamilyKind.GRAPHICS,
                )
                doc.apply_styles(style_modify)

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply color to a style
----------------------

Before applying Style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _f335230c-b84f-4beb-962d-59c8db3561e0:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f335230c-b84f-4beb-962d-59c8db3561e0
        :alt: Draw dialog Area Color style default
        :figclass: align-center
        :width: 450px

        Draw dialog Area Color style default

Apply style
^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_modify = FillColor(
            color=StandardColor.LIME_LIGHT2,
            style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
            style_family=DrawStyleFamilyKind.GRAPHICS,
        )
        doc.apply_styles(style_modify)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


After applying style
^^^^^^^^^^^^^^^^^^^^

Dialog after applying style.

.. cssclass:: screen_shot

    .. _1af864bc-5ec4-4b10-91bf-238f39818a51:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1af864bc-5ec4-4b10-91bf-238f39818a51
        :alt: Draw dialog Area Color style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Area Color style changed


Shape after applying style.

.. cssclass:: screen_shot

    .. _3f2f80c2-8231-4dfd-87b7-1c6f5ec31cc9:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/3f2f80c2-8231-4dfd-87b7-1c6f5ec31cc9
        :alt: Shape after Style applied
        :figclass: align-center

        Shape after Style applied

Getting the area image from a style
-----------------------------------

We can get the area image from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = FillColor.from_style(
            doc=doc.component,
            style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
            style_family=DrawStyleFamilyKind.GRAPHICS,
        )
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`ooodev.format.draw.modify.area.Color`
