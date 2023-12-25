.. _help_draw_format_modify_area_gradient:

Write Modify Draw Area Gradient
===============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.area.Gradient` class is used to modify the values seen in :numref:`8722a949-cef8-4e7e-90cc-0187cf9b19bf` of a style.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 25, 26, 27, 28, 29, 30

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.modify.area import Gradient, PresetGradientKind
        from ooodev.format.draw.modify.area import FamilyGraphics, DrawStyleFamilyKind


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
                style_modify = Gradient.from_preset(
                    preset=PresetGradientKind.MAHOGANY,
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

Apply gradient to a style
-------------------------

Before applying Style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _8722a949-cef8-4e7e-90cc-0187cf9b19bf:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/8722a949-cef8-4e7e-90cc-0187cf9b19bf
        :alt: Draw dialog Area Gradient style default
        :figclass: align-center
        :width: 450px

        Draw dialog Area Gradient style default

Apply style
^^^^^^^^^^^

The gradient can be loaded from a preset using the :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` class as a lookup.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_modify = Gradient.from_preset(
            preset=PresetGradientKind.MAHOGANY,
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

    .. _2e06e576-82e8-4b09-9bdd-12b3b0eacf4c:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2e06e576-82e8-4b09-9bdd-12b3b0eacf4c
        :alt: Draw dialog Area Gradient style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Area Gradient style changed


Shape after applying style.

.. cssclass:: screen_shot

    .. _a956eb5e-84c0-4651-9de0-5d2b7819cb6d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a956eb5e-84c0-4651-9de0-5d2b7819cb6d
        :alt: Shape after Style applied
        :figclass: align-center

        Shape after Style applied

Getting the area image from a style
-----------------------------------

We can get the area image from the document.

.. tabs::

    .. code-tab:: python

        # ... other code
        f_style = Gradient.from_style(
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
        - :py:class:`ooodev.format.draw.modify.area.Gradient`
        - :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind`
