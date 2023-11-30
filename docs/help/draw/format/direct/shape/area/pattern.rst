.. _help_draw_format_direct_shape_pattern:

Draw Direct Shape Area Pattern
===============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.area.Pattern` class is used to modify the values seen in :numref:`779eb709-8c21-441f-9514-29f1d0b328db` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.format.draw.direct.area import Pattern as ShapePattern
        from ooodev.format.draw.direct.area import PresetPatternKind
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
                style = ShapePattern.from_preset(preset=PresetPatternKind.SHINGLE)
                style.apply(rec)

                f_style = ShapePattern.from_obj(rec)
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

    .. _779eb709-8c21-441f-9514-29f1d0b328db:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/779eb709-8c21-441f-9514-29f1d0b328db
        :alt: Area Pattern dialog
        :figclass: align-center
        :width: 450px

        Area Pattern dialog

Add a pattern to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a fill color to the shape is done by using the ``ShapePattern`` class.
The ``ShapePattern`` class has a method ``from_preset()`` takes a ``preset`` as a parameter.
The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` class is used to get the preset of the pattern.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.area import Pattern as ShapePattern
        from ooodev.format.draw.direct.area import PresetPatternKind
        # ... other code

        rec = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
        style = ShapePattern.from_preset(preset=PresetPatternKind.SHINGLE)
        style.apply(rec)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape color can be seen in :numref:`d29db3a8-752e-43d7-a884-7b9f7d9b3aa8`.

.. cssclass:: screen_shot

    .. _d29db3a8-752e-43d7-a884-7b9f7d9b3aa8:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d29db3a8-752e-43d7-a884-7b9f7d9b3aa8
        :alt: Shape with pattern
        :figclass: align-center

        Shape with pattern

Get Shape Pattern
^^^^^^^^^^^^^^^^^^

We can get the color of the shape by using the ``ShapePattern.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.area import Pattern as ShapePattern
        # ... other code

        # get the pattern from the shape
        f_style = ShapePattern.from_obj(rec)
        assert f_style

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_direct_shape_pattern`
        - :py:class:`ooodev.format.draw.direct.area.Pattern`
