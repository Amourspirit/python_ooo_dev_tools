.. _help_draw_format_direct_shape_line_shadow:

Draw Direct Shape Line Shadow
=============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.line.Shadow` class is used to modify the values seen in :numref:`e94adcd9-ff48-472f-92fc-23ed05081e8d` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.direct.line import Shadow, ShadowLocationKind
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

                line = slide.draw_line(x1=x, y1=y, x2=x + width, y2=y + height)
                style = Shadow(
                    use_shadow=True,
                    location=ShadowLocationKind.TOP_RIGHT,
                    color=StandardColor.GRAY,
                    distance=2,
                    blur=1,
                    transparency=50,
                )
                style.apply(line.component)

                f_style = Shadow.from_obj(line.component)
                assert f_style is not None
                assert f_style.prop_use_shadow is True
                assert f_style.prop_use_shadow is True
                assert f_style.prop_location == ShadowLocationKind.TOP_RIGHT

                Lo.delay(1_000)
                doc.close_doc()
            return 0

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _e94adcd9-ff48-472f-92fc-23ed05081e8d:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/e94adcd9-ff48-472f-92fc-23ed05081e8d
        :alt: Shape Line Shadow Dialog
        :figclass: align-center
        :width: 450px

        Shape Line Shadow Dialog

Add a Line Shadow to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a Line Shadow to the shape is done by using the ``Shadow`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.line import Shadow, ShadowLocationKind
        from ooodev.utils.color import StandardColor
        # ... other code

        line = slide.draw_line(x1=x, y1=y, x2=x + width, y2=y + height)
        style = Shadow(
            use_shadow=True,
            location=ShadowLocationKind.TOP_RIGHT,
            color=StandardColor.GRAY,
            distance=2,
            blur=1,
            transparency=50,
        )
        style.apply(line.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape line shadow can be seen in :numref:`5acfab4f-b445-4f3c-b26b-64479f356512`.

.. cssclass:: screen_shot

    .. _5acfab4f-b445-4f3c-b26b-64479f356512:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/5acfab4f-b445-4f3c-b26b-64479f356512
        :alt: Shape with line shadow
        :figclass: align-center
        :width: 450px

        Shape with line shadow

Get Shape Line Shadow
^^^^^^^^^^^^^^^^^^^^^

We can get the line shadow of the shape by using the ``Shadow.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.line import Shadow
        # ... other code

        # get the shadow from the shape
        f_style = Shadow.from_obj(line.component)
        assert f_style is not None
        assert f_style.prop_use_shadow is True
        assert f_style.prop_use_shadow is True
        assert f_style.prop_location == ShadowLocationKind.TOP_RIGHT

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.line.Shadow`
