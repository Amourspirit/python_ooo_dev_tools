.. _help_draw_format_direct_transparency_transparency:

Draw Direct Shape Area Transparency Transparency
================================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.transparency.Transparency` class is used to modify the values seen in :numref:`f2b69052-dbe1-4ce9-8023-dc417774be04` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.format import Styler
        from ooodev.format.draw.direct.area import Color as ShapeColor
        from ooodev.format.draw.direct.transparency import Transparency as ShapeTransparency
        from ooodev.office.draw import Draw
        from ooodev.utils.color import StandardColor
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo


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
                style = ShapeTransparency(60)
                Styler.apply(rect, color_style, style)
                # style.apply(rect)

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

.. cssclass:: screen_shot

    .. _f2b69052-dbe1-4ce9-8023-dc417774be04:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f2b69052-dbe1-4ce9-8023-dc417774be04
        :alt: Area Transparency dialog
        :figclass: align-center
        :width: 450px

        Area Pattern dialog

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

The results of the setting the shape style can be seen in :numref:`2ee66b49-fce5-4f9c-b123-c42d490f5dcc`.

.. cssclass:: screen_shot

    .. _2ee66b49-fce5-4f9c-b123-c42d490f5dcc:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2ee66b49-fce5-4f9c-b123-c42d490f5dcc
        :alt: Shape with pattern
        :figclass: align-center

        Shape with pattern

Get Shape Transparency
^^^^^^^^^^^^^^^^^^^^^^

We can get the style of the shape by using the ``ShapeTransparency.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.area import Pattern as ShapeTransparency
        # ... other code

        # get the style from the shape
        f_style = ShapeTransparency.from_obj(rect)
        assert f_style

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_direct_shape_transparency_transparency`
        - :py:class:`ooodev.format.draw.direct.area.Pattern`
