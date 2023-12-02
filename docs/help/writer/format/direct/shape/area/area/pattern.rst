.. _help_writer_format_direct_shape_pattern:

Write Direct Shape Area Pattern
===============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.shape.area.Pattern` class is used to modify the values seen in :numref:`779eb709-8c21-441f-9514-29f1d0b328db` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.format.writer.direct.shape.area import Pattern as ShapePattern
        from ooodev.format.writer.direct.shape.area import PresetPatternKind
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
                style = ShapePattern.from_preset(preset=PresetPatternKind.SHINGLE)
                style.apply(rect)
                page.add(rect)

                f_style = ShapePattern.from_obj(rect)
                assert f_style

                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0

        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Add a pattern to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a fill pattern to the shape is done by using the ``ShapePattern`` class.
The ``ShapePattern`` class has a method ``from_preset()`` takes a ``preset`` as a parameter.
The :py:class:`~ooodev.format.inner.preset.preset_pattern.PresetPatternKind` class is used to get the preset of the pattern.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.writer.direct.shape.area import Pattern as ShapePattern
        from ooodev.format.writer.direct.shape.area import PresetPatternKind
        # ... other code

        rect = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
        style = ShapePattern.from_preset(preset=PresetPatternKind.SHINGLE)
        style.apply(rect)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape fill pattern can be seen in :numref:`68ccad1b-cc97-4311-a2f2-fdf067d37378`.

.. cssclass:: screen_shot

    .. _68ccad1b-cc97-4311-a2f2-fdf067d37378:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/68ccad1b-cc97-4311-a2f2-fdf067d37378
        :alt: Shape with pattern
        :figclass: align-center

        Shape with pattern

Get Shape Pattern
^^^^^^^^^^^^^^^^^^

We can get the fill pattern of the shape by using the ``ShapePattern.from_obj()`` method.

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

        - :ref:`help_draw_format_direct_shape_area_pattern`
        - :py:class:`ooodev.format.writer.direct.shape.area.Pattern`
