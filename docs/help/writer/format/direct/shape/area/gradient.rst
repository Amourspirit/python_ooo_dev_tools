.. _help_writer_format_direct_shape_gradient:

Write Direct Shape Area Gradient
================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.shape.area.Gradient` class is used to modify the values seen in :numref:`5ee6bb93-284f-4e54-a35b-5682d6e517b5` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        import uno

        from ooodev.format.writer.direct.shape.area import Gradient as ShapeGradient
        from ooodev.format.writer.direct.shape.area import PresetGradientKind
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.office.write import Write
        from ooodev.office.draw import Draw


        def main() -> int:
            """Main Entry Point"""

            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                page = Write.get_draw_page(doc)
                rect = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
                style = ShapeGradient.from_preset(preset=PresetGradientKind.DEEP_OCEAN)
                style.apply(rect)
                page.add(rect)

                f_style = ShapeGradient.from_obj(rect)
                assert f_style.prop_inner == style.prop_inner

                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0

        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Fill Shape Gradient
-------------------

Add a fill color to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a fill color to the shape is done by using the ``ShapeGradient`` class.
The ``ShapeGradient`` class has a method ``from_preset()`` takes a ``preset`` as a parameter.
The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` class is used to get the preset of the gradient.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.writer.direct.shape.area import Color as ShapeColor
        # ... other code

        style = ShapeGradient.from_preset(preset=PresetGradientKind.DEEP_OCEAN)
        style.apply(rect)
        page.add(rect)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape color can be seen in :numref:`8b1e7d3c-d146-4f24-b730-f574cdec999b`.

.. cssclass:: screen_shot

    .. _8b1e7d3c-d146-4f24-b730-f574cdec999b:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/8b1e7d3c-d146-4f24-b730-f574cdec999b
        :alt: Shape with Gradient color
        :figclass: align-center

        Shape with Gradient color


Get Shape Gradient
^^^^^^^^^^^^^^^^^^

We can get the color of the shape by using the ``ShapeColor.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.writer.direct.shape.area import Color as ShapeColor
        # ... other code

        # get the gradient from the shape
        f_style = ShapeGradient.from_obj(rect)
        # assert the color is the same
        assert f_style.prop_inner == style.prop_inner

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_draw_format_direct_shape_gradient`
        - :py:class:`ooodev.format.writer.direct.shape.area.Gradient`