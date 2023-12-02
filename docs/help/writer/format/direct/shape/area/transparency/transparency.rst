.. _help_writer_format_direct_shape_transparency_transparency:

Write Direct Shape Area Transparency Transparency
=================================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.shape.transparency.Transparency` class is used to modify the values seen in :numref:`f2b69052-dbe1-4ce9-8023-dc417774be04` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.format import Styler
        from ooodev.format.draw.direct.area import Color as ShapeColor
        from ooodev.format.writer.direct.shape.transparency import Transparency as ShapeTransparency
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.office.write import Write
        from ooodev.office.draw import Draw


        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                page = Write.get_draw_page(doc)
                rect = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
                color_style = ShapeColor(StandardColor.RED)
                style = ShapeTransparency(60)
                Styler.apply(rect, color_style, style)
                page.add(rect)

                f_style = ShapeTransparency.from_obj(rect)
                assert f_style

                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Add a transparency to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a transparency to the shape is done by using the ``ShapeTransparency`` class.

.. tabs::

    .. code-tab:: python

        from ooodev.format import Styler
        from ooodev.format.draw.direct.transparency import Transparency as ShapeTransparency
        # ... other code

        rect = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
        color_style = ShapeColor(StandardColor.RED)
        style = ShapeTransparency(60)
        Styler.apply(rect, color_style, style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape style can be seen in :numref:`2b8b7fe9-26d4-482e-9eca-1289070d2b37`.

.. cssclass:: screen_shot

    .. _2b8b7fe9-26d4-482e-9eca-1289070d2b37:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2b8b7fe9-26d4-482e-9eca-1289070d2b37
        :alt: Shape with pattern
        :figclass: align-center

        Shape with pattern

Get Shape Transparency
^^^^^^^^^^^^^^^^^^^^^^

We can get the style of the shape by using the ``ShapeTransparency.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.shape.transparency import Transparency as ShapeTransparency
        # ... other code

        # get the style from the shape
        f_style = ShapeTransparency.from_obj(rect)
        assert f_style

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_draw_format_direct_transparency_transparency`
        - :py:class:`ooodev.format.draw.direct.area.Pattern`
