.. _help_draw_format_direct_shape_shadow:

Draw Direct Shape Shadow
========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.shadow.Shadow` class is used to modify the values seen in :numref:`72c5790d-d5f4-4d2e-9965-da792b5e0601` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.format.draw.direct.shadow import Shadow as ShapeShadow
        from ooodev.format.draw.direct.shadow import ShadowLocationKind
        from ooodev.office.draw import Draw
        from ooodev.utils.color import StandardColor
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

                rect = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
                style = ShapeShadow(
                    use_shadow=True,
                    location=ShadowLocationKind.BOTTOM_RIGHT,
                    blur=3,
                    color=StandardColor.GRAY_LIGHT2
                )
                style.apply(rect)

                f_style = ShapeShadow.from_obj(rect)
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

    .. _72c5790d-d5f4-4d2e-9965-da792b5e0601:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/72c5790d-d5f4-4d2e-9965-da792b5e0601
        :alt: Area Pattern dialog
        :figclass: align-center
        :width: 450px

        Area Pattern dialog

Add a shadow to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a shadow to the shape is done by using the ``ShapeShadow`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.shadow import Shadow as ShapeShadow
        from ooodev.format.draw.direct.shadow import ShadowLocationKind
        # ... other code

        rect = Draw.draw_rectangle(slide=slide, x=x, y=y, width=width, height=height)
        style = ShapeShadow(
            use_shadow=True,
            location=ShadowLocationKind.BOTTOM_RIGHT,
            blur=3,
            color=StandardColor.GRAY_LIGHT2
        )
        style.apply(rect)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape shadow can be seen in :numref:`927d10fd-f71f-4a51-bd72-ecb4531ca072`.

.. cssclass:: screen_shot

    .. _927d10fd-f71f-4a51-bd72-ecb4531ca072:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/927d10fd-f71f-4a51-bd72-ecb4531ca072
        :alt: Shape with shadow
        :figclass: align-center

        Shape with shadow

Get Shape Shadow
^^^^^^^^^^^^^^^^

We can get the shadow of the shape by using the ``ShapeShadow.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.shadow import Shadow as ShapeShadow
        # ... other code

        # get the shadow from the shape
        f_style = ShapeShadow.from_obj(rect)
        assert f_style

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_direct_shape_shadow`
        - :py:class:`ooodev.format.draw.direct.area.Pattern`
