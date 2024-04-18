.. _help_writer_format_direct_shape_shadow:

Write Direct Shape Shadow
=========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.shape.shadow.Shadow` class is used to modify the values seen in :numref:`72c5790d-d5f4-4d2e-9965-da792b5e0601` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.format.writer.direct.shape.shadow import Shadow as ShapeShadow
        from ooodev.format.writer.direct.shape.shadow import ShadowLocationKind
        from ooodev.utils.color import StandardColor
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo
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
                style = ShapeShadow(
                    use_shadow=True,
                    location=ShadowLocationKind.BOTTOM_RIGHT,
                    blur=3,
                    color=StandardColor.GRAY_LIGHT2
                )
                style.apply(rect)
                page.add(rect)

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

Add a shadow to the shape
^^^^^^^^^^^^^^^^^^^^^^^^^

Adding a shadow to the shape is done by using the ``ShapeShadow`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.writer.direct.shape.shadow import Shadow as ShapeShadow
        from ooodev.format.writer.direct.shape.shadow import ShadowLocationKind
        # ... other code

        page = Write.get_draw_page(doc)
        rect = Draw.draw_rectangle(slide=page, x=10, y=10, width=100, height=100)
        style = ShapeShadow(
            use_shadow=True,
            location=ShadowLocationKind.BOTTOM_RIGHT,
            blur=3,
            color=StandardColor.GRAY_LIGHT2
        )
        style.apply(rect)
        page.add(rect)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape shadow can be seen in :numref:`cb08eee5-331e-46b8-805c-9f013d19e819`.

.. cssclass:: screen_shot

    .. _cb08eee5-331e-46b8-805c-9f013d19e819:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/cb08eee5-331e-46b8-805c-9f013d19e819
        :alt: Shape with shadow
        :figclass: align-center

        Shape with shadow

Get Shape Shadow
^^^^^^^^^^^^^^^^

We can get the shadow of the shape by using the ``ShapeShadow.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.shape.shadow import Shadow as ShapeShadow
        # ... other code

        # get the shadow from the shape
        f_style = ShapeShadow.from_obj(rect)
        assert f_style

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_draw_format_direct_shape_shadow`
        - :py:class:`ooodev.format.draw.direct.area.Pattern`
