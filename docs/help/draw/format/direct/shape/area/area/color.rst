.. _help_draw_format_direct_shape_area_color:

Draw Direct Shape Area Color
============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.area.Color` class is used to modify the values seen in :numref:`5ee6bb93-284f-4e54-a35b-5682d6e517b5` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.format.draw.direct.area import Color as ShapeColor
        from ooodev.utils.color import StandardColor
        from ooodev.office.draw import Draw
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

                rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
                style = ShapeColor(color=StandardColor.GREEN_LIGHT2)
                style.apply(rec)

                f_style = ShapeColor.from_obj(rec)
                assert f_style.prop_color == style.prop_color

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _5ee6bb93-284f-4e54-a35b-5682d6e517b5:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/5ee6bb93-284f-4e54-a35b-5682d6e517b5
        :alt: Area color dialog
        :figclass: align-center
        :width: 450px

        Area color dialog

Fill Shape Color
----------------

Add a fill color to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a fill color to the shape is done by using the ``ShapeColor`` class.
The ``ShapeColor`` class takes a ``color`` as a parameter.
The :py:class:`~ooodev.utils.color.StandardColor` class is used to set the color of the shape.


.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.area import Color as ShapeColor
        # ... other code

        rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
        style = ShapeColor(color=StandardColor.GREEN_LIGHT2)
        style.apply(rec)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape color can be seen in :numref:`2c8a395c-9495-49af-90cc-30d4b4a5e4d2`.

.. cssclass:: screen_shot

    .. _2c8a395c-9495-49af-90cc-30d4b4a5e4d2:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2c8a395c-9495-49af-90cc-30d4b4a5e4d2
        :alt: Shape with fill color
        :figclass: align-center

        Shape with fill color


Get Shape Color
^^^^^^^^^^^^^^^

We can get the color of the shape by using the ``ShapeColor.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.area import Color as ShapeColor
        # ... other code

        # get the color from the shape
        f_style = ShapeColor.from_obj(rec)
        # assert the color is the same
        assert f_style.prop_color == style.prop_color

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_direct_shape_color`
        - :py:class:`ooodev.format.draw.direct.area.Color`