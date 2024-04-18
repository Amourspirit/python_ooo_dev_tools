.. _help_draw_format_direct_shape_area_gradient:

Draw Direct Shape Area Gradient
===============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.area.Gradient` class is used to modify the values seen in :numref:`4fdbcc5e-6621-4e89-bb20-3ad4974a3e4b` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.format.draw.direct.area import Gradient as ShapeGradient
        from ooodev.format.draw.direct.area import PresetGradientKind
        from ooodev.office.draw import Draw
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

                rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
                style = ShapeGradient.from_preset(preset=PresetGradientKind.DEEP_OCEAN)
                style.apply(rec)

                f_style = ShapeGradient.from_obj(rec)
                assert f_style.prop_inner == style.prop_inner

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _4fdbcc5e-6621-4e89-bb20-3ad4974a3e4b:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4fdbcc5e-6621-4e89-bb20-3ad4974a3e4b
        :alt: Area Gradient dialog
        :figclass: align-center
        :width: 450px

        Area Gradient dialog

Add a gradient to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a gradient to the shape is done by using the ``ShapeGradient`` class.
The ``ShapeGradient`` class has a method ``from_preset()`` takes a ``preset`` as a parameter.
The :py:class:`~ooodev.format.inner.preset.preset_gradient.PresetGradientKind` class is used to get the preset of the gradient.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.area import Gradient as ShapeGradient
        from ooodev.format.draw.direct.area import PresetGradientKind
        # ... other code

        rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
        style = ShapeGradient.from_preset(preset=PresetGradientKind.DEEP_OCEAN)
        style.apply(rec)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape gradient can be seen in :numref:`b4dd891c-8512-4285-bdf6-671e4ffdc97d`.

.. cssclass:: screen_shot

    .. _b4dd891c-8512-4285-bdf6-671e4ffdc97d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/b4dd891c-8512-4285-bdf6-671e4ffdc97d
        :alt: Shape with gradient
        :figclass: align-center

        Shape with gradient

Get Shape Gradient
^^^^^^^^^^^^^^^^^^

We can get the gradient of the shape by using the ``ShapeGradient.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.area import Gradient as ShapeGradient
        # ... other code

        # get the gradient from the shape
        f_style = ShapeGradient.from_obj(rec)
        # assert the color is the same
        assert f_style.prop_inner == style.prop_inner

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_direct_shape_gradient`
        - :py:class:`ooodev.format.draw.direct.area.Gradient`
