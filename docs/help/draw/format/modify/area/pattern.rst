.. _help_draw_format_modify_area_pattern:

Write Modify Draw Area Pattern
==============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.area.Pattern` class is used to modify the values seen in :numref:`eed969b9-f125-4638-8d07-d93668e42029` of a style.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 25, 26, 27, 28, 29

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.modify.area import Pattern, PresetPatternKind
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
                style_modify = Pattern.from_preset(
                    preset=PresetPatternKind.SHINGLE,
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

    .. _eed969b9-f125-4638-8d07-d93668e42029:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/eed969b9-f125-4638-8d07-d93668e42029
        :alt: Draw dialog Area Pattern style default
        :figclass: align-center
        :width: 450px

        Draw dialog Area Pattern style default

Apply style
^^^^^^^^^^^

The gradient can be loaded from a preset using the :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` class as a lookup.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_modify = Pattern.from_preset(
            preset=PresetPatternKind.SHINGLE,
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

    .. _5a815341-75bb-400b-b266-0611ef54f5a8:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/5a815341-75bb-400b-b266-0611ef54f5a8
        :alt: Draw dialog Area Pattern style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Area Pattern style changed


Shape after applying style.

.. cssclass:: screen_shot

    .. _9d9f2545-d6df-4a7c-bf1c-83ee6d4df9f5:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9d9f2545-d6df-4a7c-bf1c-83ee6d4df9f5
        :alt: Shape after Style applied
        :figclass: align-center

        Shape after Style applied

Getting the area pattern from a style
-------------------------------------

We can get the area pattern from the document.

.. tabs::

    .. code-tab:: python

        # ... other code
        f_style = Pattern.from_style(
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
        - :py:class:`ooodev.format.draw.modify.area.Pattern`
        - :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind`
