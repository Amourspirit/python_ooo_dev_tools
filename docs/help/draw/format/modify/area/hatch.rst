.. _help_draw_format_modify_area_hatch:

Write Modify Draw Area Hatch
============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.area.Hatch` class is used to modify the values seen in :numref:`a5a55e3b-65ca-400c-a245-b1ba655507e1` of a style.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 25, 26, 27, 28, 29

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.modify.area import Hatch, PresetHatchKind
        from ooodev.format.draw.modify.area import FamilyGraphics, DrawStyleFamilyKind

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
                style_modify = Hatch.from_preset(
                    preset=PresetHatchKind.GREEN_30_DEGREES,
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

    .. _a5a55e3b-65ca-400c-a245-b1ba655507e1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a5a55e3b-65ca-400c-a245-b1ba655507e1
        :alt: Draw dialog Area Hatch style default
        :figclass: align-center
        :width: 450px

        Draw dialog Area Hatch style default

Apply style
^^^^^^^^^^^

The gradient can be loaded from a preset using the :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` class as a lookup.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_modify = Hatch.from_preset(
            preset=PresetHatchKind.GREEN_30_DEGREES,
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

    .. _2f703ba4-f5bc-4e20-a328-043aaae746b3:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2f703ba4-f5bc-4e20-a328-043aaae746b3
        :alt: Draw dialog Area Hatch style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Area Hatch style changed


Shape after applying style.

.. cssclass:: screen_shot

    .. _9b29a4fa-496d-47a4-9f42-9c6c7ef73173:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9b29a4fa-496d-47a4-9f42-9c6c7ef73173
        :alt: Shape after Style applied
        :figclass: align-center

        Shape after Style applied

Getting the area pattern from a style
-------------------------------------

We can get the area pattern from the document.

.. tabs::

    .. code-tab:: python

        # ... other code
        f_style = Hatch.from_style(
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
        - :py:class:`ooodev.format.draw.modify.area.Hatch`
        - :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind`
