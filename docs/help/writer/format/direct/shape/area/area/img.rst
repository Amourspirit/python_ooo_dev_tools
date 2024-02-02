.. _help_writer_format_direct_shape_image:

Draw Direct Shape Area Image
============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.shape.area.Img` class is used to modify the values seen in :numref:`0daf9c23-f6c6-41a8-9b05-2b24f2dce71b` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.format.draw.direct.area import Img as ShapeImage
        from ooodev.format.draw.direct.area import PresetImageKind
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
                style = ShapeImage.from_preset(preset=PresetImageKind.COFFEE_BEANS)
                style.apply(rec)

                f_style = ShapeImage.from_obj(rec)
                assert f_style.prop_size
                assert f_style.prop_size == style.prop_size

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Add a image to the shape
^^^^^^^^^^^^^^^^^^^^^^^^

Adding a fill image to the shape is done by using the ``ShapeImage`` class.
The ``ShapeImage`` class has a method ``from_preset()`` takes a ``preset`` as a parameter.
The :py:class:`~ooodev.format.inner.preset.preset_image.PresetImageKind` class is used to get the preset of the image.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.area import Img as ShapeImage
        from ooodev.format.draw.direct.area import PresetImageKind
        # ... other code

        page = Write.get_draw_page(doc)
        rect = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
        style = ShapeImage.from_preset(preset=PresetImageKind.COFFEE_BEANS)
        style.apply(rect)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape fill image can be seen in :numref:`471fa925-2e21-4e6b-8f65-05c893f8dd0f`.

.. cssclass:: screen_shot

    .. _471fa925-2e21-4e6b-8f65-05c893f8dd0f:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/471fa925-2e21-4e6b-8f65-05c893f8dd0f
        :alt: Shape with Image
        :figclass: align-center

        Shape with gradient

Get Shape Image
^^^^^^^^^^^^^^^

We can get the fill image of the shape by using the ``ShapeImage.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.area import Img as ShapeImage
        # ... other code

        # get the image from the shape
        f_style = ShapeImage.from_obj(rect)
        assert f_style.prop_size
        assert f_style.prop_size == style.prop_size

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_draw_format_direct_shape_area_image`
        - :py:class:`ooodev.format.writer.direct.shape.area.Img`
