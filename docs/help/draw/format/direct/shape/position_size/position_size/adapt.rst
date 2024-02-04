.. _help_draw_format_direct_shape_position_size_position_size_adapt:

Draw Direct Shape Position Size - Adapt
=======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.position_size.position_size.Adapt` class is used to modify the values seen in :numref:`c4bf44f7-94bc-4e6c-a713-11c5856bfa88` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.format.draw.direct.position_size.position_size import Adapt


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

                text = slide.draw_text(msg="Hello World", x=x, y=y, width=width, height=height)
                style = Adapt(fit_height=True, fit_width=True)
                style.apply(text.component)

                f_style = Adapt.from_obj(text.component)
                assert f_style is not None

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set Adaption of a Shape
^^^^^^^^^^^^^^^^^^^^^^^

Setting Adapt of the shape is done by using the ``Adapt`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.position_size.position_size import Adapt
        # ... other code

        text = slide.draw_text(
            msg="Hello World", x=x, y=y, width=width, height=height
        )
        style = Adapt(fit_height=True, fit_width=True)
        style.apply(text.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape adaption can be seen in :numref:`708a0b0d-17d9-409e-904b-e21330a552b6`.

.. cssclass:: screen_shot

    .. _708a0b0d-17d9-409e-904b-e21330a552b6:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/708a0b0d-17d9-409e-904b-e21330a552b6
        :alt: Text Shape with Adapt set
        :figclass: align-center
        :width: 450px

        Text Shape with Adapt set

.. note::

    Adapt only applies to text shapes.

Get Shape Adapt
^^^^^^^^^^^^^^^

We can get the Adapt properties of the shape by using the ``Adapt.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.position_size.position_size import Adapt
        # ... other code

        # get the properties from the shape
        f_style = Adapt.from_obj(text.component)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.position_size.position_size.Adapt`
