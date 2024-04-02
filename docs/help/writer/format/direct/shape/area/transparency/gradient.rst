.. _help_writer_format_direct_shape_transparency_gradient:

Write Direct Shape Area Transparency Gradient
=============================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.shape.transparency.Gradient` class is used to modify the values seen in :numref:`aab455a0-c7e7-42f1-a7ea-9b4210732ec9` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.format import Styler
        from ooodev.format.draw.direct.area import Color as ShapeColor
        from ooodev.format.draw.direct.transparency import Gradient as ShapeGradient
        from ooodev.format.draw.direct.transparency import GradientStyle, IntensityRange
        from ooodev.office.draw import Draw
        from ooodev.utils.color import StandardColor
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo


        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                page = Write.get_draw_page(doc)
                rect = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
                color_style = ShapeColor(StandardColor.RED)
                style = ShapeGradient(
                    style=GradientStyle.LINEAR, angle=30, grad_intensity=IntensityRange(0, 100)
                )
                Styler.apply(rect, color_style, style)
                page.add(rect)

                f_style = ShapeGradient.from_obj(rect)
                assert f_style

                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Add a gradient transparency to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a transparency gradient to the shape is done by using the ``ShapeGradient`` class.

.. tabs::

    .. code-tab:: python

        from ooodev.format import Styler
        from ooodev.format.draw.direct.area import Color as ShapeColor
        from ooodev.format.writer.direct.shape.transparency import Gradient as ShapeGradient
        from ooodev.format.writer.direct.shape.transparency import GradientStyle, IntensityRange
        # ... other code

        page = Write.get_draw_page(doc)
        rect = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
        color_style = ShapeColor(StandardColor.RED)
        style = ShapeGradient(
            style=GradientStyle.LINEAR, angle=30, grad_intensity=IntensityRange(0, 100)
        )
        Styler.apply(rect, color_style, style)
        page.add(rect)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape style can be seen in :numref:`52459d3d-4527-49f1-9631-40ac53fe2eb8`.

.. cssclass:: screen_shot

    .. _52459d3d-4527-49f1-9631-40ac53fe2eb8:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/52459d3d-4527-49f1-9631-40ac53fe2eb8
        :alt: Shape with Transparency Gradient
        :figclass: align-center

        Shape with Transparency Gradient

Get Shape Transparency Gradient
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the transparency gradient of the shape by using the ``ShapeGradient.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.shape.transparency import Gradient as ShapeGradient
        # ... other code

        # get the style from the shape
        f_style = ShapeGradient.from_obj(rect)
        assert f_style

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_draw_format_direct_transparency_gradient`
        - :py:class:`ooodev.format.draw.direct.area.Pattern`
