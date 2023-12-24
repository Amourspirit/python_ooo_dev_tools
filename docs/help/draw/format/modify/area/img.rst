.. _help_draw_format_modify_area_image:

Write Modify Draw Area Image
============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.area.Img` class is used to modify the values seen in :numref:`76aab8df-4dc9-42f6-b677-9b19d3ce7501` of a paragraph style.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.modify.area import Img as FillImg
        from ooodev.format.draw.modify.area import PresetImageKind

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
                style_modify = FillImg.from_preset(preset=PresetImageKind.POOL)
                doc.apply_styles(style_modify)

                Lo.delay(1_000)
                doc.close_doc()
            return 0

        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply image to a style
----------------------

Before applying Style
^^^^^^^^^^^^^^^^^^^^^

.. cssclass:: screen_shot

    .. _76aab8df-4dc9-42f6-b677-9b19d3ce7501:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/76aab8df-4dc9-42f6-b677-9b19d3ce7501
        :alt: Draw dialog Area Image style default
        :figclass: align-center
        :width: 450px

        Draw dialog Area Image style default

Apply style
^^^^^^^^^^^

The image can be loaded from a preset using the :py:class:`ooodev.format.inner.preset.preset_image.PresetImageKind` class as a lookup.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_modify = FillImg.from_preset(preset=PresetImageKind.POOL)
        doc.apply_styles(style_modify)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


After applying style
^^^^^^^^^^^^^^^^^^^^

Dialog after applying style.

.. cssclass:: screen_shot

    .. _8ea541ab-ffea-451c-bf56-93fe00ca99eb:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/8ea541ab-ffea-451c-bf56-93fe00ca99eb
        :alt: Draw dialog Area Image style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Area Image style changed


Shape after applying style.

.. cssclass:: screen_shot

    .. _9ecb81d6-66b9-4499-add2-3ac48b95dd8f:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9ecb81d6-66b9-4499-add2-3ac48b95dd8f
        :alt: Shape after Style applied
        :figclass: align-center

        Shape after Style applied

Getting the area image from a style
-----------------------------------

We can get the area image from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = FillImg.from_style(
            doc=doc.component,
            style_name=style_modify.prop_style_name,
            style_family=style_modify.prop_style_family_name,
        )
        assert f_style.prop_style_name == style_modify.prop_style_name

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`ooodev.format.draw.modify.area.Img`
        - :py:class:`ooodev.format.inner.preset.preset_image.PresetImageKind`
