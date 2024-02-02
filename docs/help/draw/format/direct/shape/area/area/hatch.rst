.. _help_draw_format_direct_shape_area_hatch:

Draw Direct Shape Area Hatch
============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.area.Hatch` class is used to modify the values seen in :numref:`bc655b90-8c12-4ac3-a768-448a2b07e5c3` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.format.draw.direct.area import Hatch as ShapeHatch
        from ooodev.format.draw.direct.area import PresetHatchKind
        from ooodev.office.draw import Draw
        from ooodev.utils.gui import GUI
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
                style = ShapeHatch.from_preset(preset=PresetHatchKind.GREEN_30_DEGREES)
                style.apply(rec)

                f_style = ShapeHatch.from_obj(rec)
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

    .. _bc655b90-8c12-4ac3-a768-448a2b07e5c3:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/bc655b90-8c12-4ac3-a768-448a2b07e5c3
        :alt: Area Hatch dialog
        :figclass: align-center
        :width: 450px

        Area Hatch dialog

Add a hatch to the shape
^^^^^^^^^^^^^^^^^^^^^^^^

Adding a hatch for the shape is done by using the ``ShapeGradient`` class.
The ``ShapeGradient`` class has a method ``from_preset()`` takes a ``preset`` as a parameter.
The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` class is used to get the preset of the hatch.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.area import Hatch as ShapeHatch
        from ooodev.format.draw.direct.area import PresetHatchKind
        # ... other code

        rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
        style = ShapeHatch.from_preset(preset=PresetHatchKind.GREEN_30_DEGREES)
        style.apply(rec)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape hatch can be seen in :numref:`4965571a-4918-4b64-8f8a-87203f1d7b3a`.

.. cssclass:: screen_shot

    .. _4965571a-4918-4b64-8f8a-87203f1d7b3a:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4965571a-4918-4b64-8f8a-87203f1d7b3a
        :alt: Shape with hatch
        :figclass: align-center

        Shape with hatch

Get Shape Hatch
^^^^^^^^^^^^^^^

We can get the hatch of the shape by using the ``ShapeHatch.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.area import Hatch as ShapeHatch
        # ... other code

        # get the hatch from the shape
        f_style = ShapeHatch.from_obj(rec)
        assert f_style

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_direct_shape_hatch`
        - :py:class:`ooodev.format.draw.direct.area.Hatch`
