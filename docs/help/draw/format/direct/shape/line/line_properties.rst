.. _help_draw_format_direct_shape_line_properties:

Draw Direct Shape Line Properties
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.line.LineProperties` class is used to modify the values seen in :numref:`bb764719-da32-4497-874c-a54f7c9f7aaa` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.direct.line import LineProperties, BorderLineKind
        from ooodev.utils.color import StandardColor


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = DrawDoc(Draw.create_draw_doc())
                doc.set_visible()
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_75_PERCENT)

                slide = doc.get_slide()

                width = 36
                height = 36
                x = round(width / 2)
                y = round(height / 2)

                rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
                style = LineProperties(
                    style=BorderLineKind.CONTINUOUS,
                    color=StandardColor.RED_DARK3,
                    width=0.7,
                    transparency=22
                )
                style.apply(rect.component)

                f_style = LineProperties.from_obj(rect.component)
                assert f_style is not None
                assert f_style.prop_color == StandardColor.RED_DARK3

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _bb764719-da32-4497-874c-a54f7c9f7aaa:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/bb764719-da32-4497-874c-a54f7c9f7aaa
        :alt: Area Pattern dialog
        :figclass: align-center
        :width: 450px

        Area Pattern dialog

Add a Line Properties to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a Line Properties to the shape is done by using the ``LineProperties`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.line import LineProperties, BorderLineKind
        from ooodev.utils.color import StandardColor
        # ... other code

        rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        style = LineProperties(
            style=BorderLineKind.CONTINUOUS,
            color=StandardColor.RED_DARK3,
            width=0.7,
            transparency=22
        )
        style.apply(rect.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape line properties can be seen in :numref:`673594a5-649e-4549-8436-acbb20210c9e`.

.. cssclass:: screen_shot

    .. _673594a5-649e-4549-8436-acbb20210c9e:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/673594a5-649e-4549-8436-acbb20210c9e
        :alt: Shape with line properties
        :figclass: align-center

        Shape with line properties

Get Shape Line Properties
^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the line properties of the shape by using the ``LineProperties.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.line import LineProperties
        # ... other code

        # get the shadow from the shape
        f_style = LineProperties.from_obj(rect.component)
        assert f_style.prop_color == StandardColor.RED_DARK3

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.line.LineProperties`
