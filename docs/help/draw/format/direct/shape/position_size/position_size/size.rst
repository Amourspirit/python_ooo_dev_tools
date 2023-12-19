.. _help_draw_format_direct_shape_position_size_position_size_size:

Draw Direct Shape Position Size - Size
======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.position_size.position_size.Size` class is used to modify the values seen in :numref:`c4bf44f7-94bc-4e6c-a713-11c5856bfa88` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.direct.position_size.position_size import Size


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
                style = Size(width=50, height=50)
                style.apply(rect.component)

                f_style = Size.from_obj(rect.component)
                assert f_style is not None

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set Size of a Shape
^^^^^^^^^^^^^^^^^^^

Setting size of the shape is done by using the ``Size`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.position_size.position_size import Size
        # ... other code

        rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        style = Size(width=50, height=50)
        style.apply(rect.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape size can be seen in :numref:`b117ba51-1e5f-4962-bdd4-6dd879988451`.

.. cssclass:: screen_shot

    .. _b117ba51-1e5f-4962-bdd4-6dd879988451:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/b117ba51-1e5f-4962-bdd4-6dd879988451
        :alt: Shape with size set
        :figclass: align-center

        Shape with size set

Get Shape Size
^^^^^^^^^^^^^^

We can get the position of the shape by using the ``Size.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.position_size.position_size import Size
        # ... other code

        # get the size from the shape
        f_style = Size.from_obj(rect.component)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.position_size.position_size.Size`
