.. _help_draw_format_direct_shape_line_arrow_styles:

Draw Direct Shape Line Arrow Styles
===================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.line.ArrowStyles` class is used to modify the values seen in :numref:`936ba317-b048-4a7f-996c-ad49f62dcdec` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.direct.line import ArrowStyles, GraphicArrowStyleKind


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

                line = slide.draw_line(x1=x, y1=y, x2=x + width, y2=y + height)
                style = ArrowStyles(
                    start_line_name=GraphicArrowStyleKind.ARROW_LARGE,
                    start_line_center=True,
                    start_line_width=2.5,
                    end_line_name=GraphicArrowStyleKind.SQUARE_45,
                    end_line_center=False,
                    end_line_width=1.9,
                )
                style.apply(line.component)

                f_style = ArrowStyles.from_obj(line.component)
                assert f_style is not None
                assert f_style.prop_start_line_name == GraphicArrowStyleKind.ARROW_LARGE.value

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _936ba317-b048-4a7f-996c-ad49f62dcdec:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/936ba317-b048-4a7f-996c-ad49f62dcdec
        :alt: Shape Line Arrow Styles Dialog
        :figclass: align-center
        :width: 450px

        Shape Line Arrow Styles Dialog

Add a Line Properties to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a Line Arrow Styles to the shape is done by using the ``ArrowStyles`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.line import ArrowStyles, GraphicArrowStyleKind

        # ... other code

        line = slide.draw_line(x1=x, y1=y, x2=x + width, y2=y + height)
        style = ArrowStyles(
            start_line_name=GraphicArrowStyleKind.ARROW_LARGE,
            start_line_center=True,
            start_line_width=2.5,
            end_line_name=GraphicArrowStyleKind.SQUARE_45,
            end_line_center=False,
            end_line_width=1.9,
        )
        style.apply(line.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape line properties can be seen in :numref:`2bb1841f-d7af-46bf-9a0d-f76eac3ccb19`.

.. cssclass:: screen_shot

    .. _2bb1841f-d7af-46bf-9a0d-f76eac3ccb19:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2bb1841f-d7af-46bf-9a0d-f76eac3ccb19
        :alt: Shape line Arrow Styles
        :figclass: align-center
        :width: 450px

        Shape line Arrow Styles

Get Shape Line Arrow Styles
^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the line arrow styles of the shape by using the ``ArrowStyles.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.line import ArrowStyles
        # ... other code

        # get the properties from the shape
        f_style = ArrowStyles.from_obj(line.component)
        assert f_style is not None
        assert f_style.prop_start_line_name == GraphicArrowStyleKind.ARROW_LARGE.value

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.line.ArrowStyles`
