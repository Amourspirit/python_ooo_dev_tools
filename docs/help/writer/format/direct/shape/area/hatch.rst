.. _help_writer_format_direct_shape_hatch:

Write Direct Shape Area Hatch
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.shape.area.Hatch` class is used to modify the values seen in :numref:`bc655b90-8c12-4ac3-a768-448a2b07e5c3` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.writer.direct.shape.area import Hatch as ShapeHatch
        from ooodev.format.writer.direct.shape.area import PresetHatchKind
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
                style = ShapeHatch.from_preset(preset=PresetHatchKind.GREEN_30_DEGREES)
                style.apply(rect)
                page.add(rect)

                f_style = ShapeHatch.from_obj(rect)
                assert f_style

                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Fill Shape Hatch
----------------

Add a hatch fill to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a hatch to the shape is done by using the ``ShapeHatch`` class.
The ``ShapeHatch`` class has a method ``from_preset()`` takes a ``preset`` as a parameter.
The :py:class:`~ooodev.format.inner.preset.preset_hatch.PresetHatchKind` class is used to get the preset of the hatch.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.writer.direct.shape.area import Hatch as ShapeHatch
        # ... other code

        style = ShapeHatch.from_preset(preset=PresetHatchKind.GREEN_30_DEGREES)
        style.apply(rect)
        page.add(rect)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape color can be seen in :numref:`59bd525e-7e1e-49ca-8251-64a4355486c0`.

.. cssclass:: screen_shot

    .. _59bd525e-7e1e-49ca-8251-64a4355486c0:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/59bd525e-7e1e-49ca-8251-64a4355486c0
        :alt: Shape with Hatch color
        :figclass: align-center

        Shape with Hatch color


Get Shape Hatch
^^^^^^^^^^^^^^^

We can get the color of the shape by using the ``ShapeHatch.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.writer.direct.shape.area import Hatch as ShapeHatch
        # ... other code

        # get the gradient from the shape
        f_style = ShapeHatch.from_obj(rect)
        assert f_style

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_draw_format_direct_shape_hatch`
        - :py:class:`ooodev.format.writer.direct.shape.area.Hatch`