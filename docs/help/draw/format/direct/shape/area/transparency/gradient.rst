.. _help_draw_format_direct_transparency_gradient:

Draw Direct Shape Area Transparency Gradient
============================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.transparency.Gradient` class is used to modify the values seen in :numref:`aab455a0-c7e7-42f1-a7ea-9b4210732ec9` of a shape.

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
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Draw.create_draw_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)

                slide = Draw.get_slide(doc)

                width = 36
                height = 36
                x = int(width / 2)
                y = int(height / 2) + 20

                rect = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
                color_style = ShapeColor(StandardColor.RED)
                style = ShapeGradient(
                    style=GradientStyle.LINEAR, angle=30, grad_intensity=IntensityRange(0, 100)
                )
                Styler.apply(rect, color_style, style)
                # style.apply(rect)

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

.. cssclass:: screen_shot

    .. _aab455a0-c7e7-42f1-a7ea-9b4210732ec9:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/aab455a0-c7e7-42f1-a7ea-9b4210732ec9
        :alt: Area Transparency Gradient dialog
        :figclass: align-center
        :width: 450px

        Area Transparency Gradient dialog

Add a gradient transparency to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a transparency gradient to the shape is done by using the ``ShapeGradient`` class.

.. tabs::

    .. code-tab:: python

        from ooodev.format import Styler
        from ooodev.format.draw.direct.area import Color as ShapeColor
        from ooodev.format.draw.direct.transparency import Gradient as ShapeGradient
        from ooodev.format.draw.direct.transparency import GradientStyle, IntensityRange
        # ... other code

        rect = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
        color_style = ShapeColor(StandardColor.RED)
        style = ShapeGradient(
            style=GradientStyle.LINEAR, angle=30, grad_intensity=IntensityRange(0, 100)
        )
        Styler.apply(rect, color_style, style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape transparency gradient can be seen in :numref:`e8cc476e-1aa1-4999-8161-2bd4d25729ae`.

.. cssclass:: screen_shot

    .. _e8cc476e-1aa1-4999-8161-2bd4d25729ae:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/e8cc476e-1aa1-4999-8161-2bd4d25729ae
        :alt: Shape with Transparency Gradient
        :figclass: align-center

        Shape with Transparency Gradient

Get Shape Transparency Gradient
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the transparency gradient of the shape by using the ``ShapeGradient.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.transparency import Gradient as ShapeGradient
        # ... other code

        # get the style from the shape
        f_style = ShapeGradient.from_obj(rect)
        assert f_style

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_direct_shape_transparency_gradient`
        - :py:class:`ooodev.format.draw.direct.area.Pattern`
