.. _help_draw_format_direct_shape_position_size_position_size_position:

Draw Direct Shape Position Size - Position
==========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.position_size.position_size.Position` class is used to modify the values seen in :numref:`c4bf44f7-94bc-4e6c-a713-11c5856bfa88` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.direct.position_size.position_size import Position


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
                pos = Position(pos_x=38, pos_y=18)
                pos.apply(rect.component)

                pos2 = Position.from_obj(rect.component)
                assert pos2 is not None

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _c4bf44f7-94bc-4e6c-a713-11c5856bfa88:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/c4bf44f7-94bc-4e6c-a713-11c5856bfa88
        :alt: Shape Size and Position Dialog
        :figclass: align-center
        :width: 450px

        Shape Size and Position Dialog

Set Position of a Shape
^^^^^^^^^^^^^^^^^^^^^^^

Setting position of the shape is done by using the ``Position`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.position_size.position_size import Position
        # ... other code

        rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        pos = Position(pos_x=38, pos_y=18)
        pos.apply(rect.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape position can be seen in :numref:`d2de7b14-c995-44c5-8b49-9d4eec8f2156`.

.. cssclass:: screen_shot

    .. _d2de7b14-c995-44c5-8b49-9d4eec8f2156:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d2de7b14-c995-44c5-8b49-9d4eec8f2156
        :alt: Shape with position set
        :figclass: align-center

        Shape with position set

.. note::

    ``pos_x`` and ``pos_y`` are the coordinates of the shape inside the draw page borders.
    This is the same behavior as the dialog box.
    If the draw page has a border of 10mm and the shape is positioned at ``0 mm``, ``0 mm`` in the dialog box then the shape
    is actually at ``10 mm``, ``10 mm`` relative to the draw page document.

Get Shape Position
^^^^^^^^^^^^^^^^^^

We can get the position of the shape by using the ``Position.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.position_size.position_size import Position
        # ... other code

        # get the position from the shape
        pos2 = Position.from_obj(rect.component)
        assert pos2 is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.position_size.position_size.Position`
