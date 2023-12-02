.. _help_writer_format_direct_shape_color:

Write Direct Shape Area Color
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.shape.area.Color` class is used to modify the values seen in :numref:`5ee6bb93-284f-4e54-a35b-5682d6e517b5` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.writer.direct.shape.area import Color as ShapeColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.office.write import Write
        from ooodev.office.draw import Draw
        from ooodev.utils.color import StandardColor


        def main() -> int:
            """Main Entry Point"""

            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                style = ShapeColor(color=StandardColor.GREEN_LIGHT2)

                page = Write.get_draw_page(doc)
                rect = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
                style.apply(rect)
                page.add(rect)

                f_style = ShapeColor.from_obj(rect)
                assert f_style.prop_color == style.prop_color

                Lo.delay(1_000)

                Lo.close_doc(doc)


        if __name__ == "__main__":
            raise SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Fill Shape Color
----------------

Add a fill color to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a fill color to the shape is done by using the ``ShapeColor`` class.
The ``ShapeColor`` class takes a ``color`` as a parameter.
The :py:class:`~ooodev.utils.color.StandardColor` class is used to set the color of the shape.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.writer.direct.shape.area import Color as ShapeColor
        # ... other code

        style = ShapeColor(color=StandardColor.GREEN_LIGHT2)

        # get the page
        page = Write.get_draw_page(doc)
        # draw a rectangle
        rect = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
        # apply the style
        style.apply(rect)
        # add the rectangle to the page
        page.add(rect)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape color can be seen in :numref:`f5ab7604-3893-457b-8849-4b6cef1ade01`.

.. cssclass:: screen_shot

    .. _f5ab7604-3893-457b-8849-4b6cef1ade01:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/f5ab7604-3893-457b-8849-4b6cef1ade01
        :alt: Shape with fill color
        :figclass: align-center

        Shape with fill color


Get Shape Color
^^^^^^^^^^^^^^^

We can get the color of the shape by using the ``ShapeColor.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.writer.direct.shape.area import Color as ShapeColor
        # ... other code

        # get the color from the shape
        f_style = ShapeColor.from_obj(rect)
        # assert the color is the same
        assert f_style.prop_color == style.prop_color

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_draw_format_direct_shape_area_color`
        - :py:class:`ooodev.format.writer.direct.shape.area.Color`