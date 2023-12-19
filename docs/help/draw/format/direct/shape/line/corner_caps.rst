.. _help_draw_format_direct_shape_line_corner_caps:

Draw Direct Shape Corner and Caps
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.line.CornerCaps` class is used to modify the values seen in :numref:`479a3c71-3377-4c45-bb58-aaa6754085ad` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.direct.line import CornerCaps, LineJoint, LineCap


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
                style = CornerCaps(
                    corner_style=LineJoint.MIDDLE,
                    cap_style=LineCap.SQUARE,
                )
                style.apply(rect.component)

                f_style = CornerCaps.from_obj(rect.component)
                assert f_style is not None
                assert f_style.prop_corner_style == LineJoint.MIDDLE
                assert f_style.prop_cap_style == LineCap.SQUARE

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _479a3c71-3377-4c45-bb58-aaa6754085ad:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/479a3c71-3377-4c45-bb58-aaa6754085ad
        :alt: Shape Corner and Caps Dialog
        :figclass: align-center
        :width: 450px

        Shape Corner and Caps Dialog

Add a Corner and Caps style to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a Corner and Caps style to the shape is done by using the ``CornerCaps`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.line import CornerCaps, LineJoint, LineCap
        # ... other code

        rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        style = CornerCaps(
            corner_style=LineJoint.MIDDLE,
            cap_style=LineCap.SQUARE,
        )
        style.apply(rect.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape corner and cap style can be seen in :numref:`e3127b2f-43e0-4f83-9786-b72c5f06e0f9`.

.. cssclass:: screen_shot

    .. _e3127b2f-43e0-4f83-9786-b72c5f06e0f9:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/e3127b2f-43e0-4f83-9786-b72c5f06e0f9
        :alt: Shape Corner and Caps Dialog
        :figclass: align-center
        :width: 450px

        Shape Corner and Caps Dialog

Get Shape Corner and Caps style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the corner and caps style of the shape by using the ``CornerCaps.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.line import CornerCaps
        # ... other code

        # get the properties from the shape
        f_style = CornerCaps.from_obj(rect.component)
        assert f_style.prop_corner_style == LineJoint.MIDDLE
        assert f_style.prop_cap_style == LineCap.SQUARE

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.line.CornerCaps`
