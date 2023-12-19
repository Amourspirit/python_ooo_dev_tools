.. _help_draw_format_direct_shape_position_size_position_size_protect:

Draw Direct Shape Position Size - Protect
=========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.position_size.position_size.Protect` class is used to modify the values seen in :numref:`c4bf44f7-94bc-4e6c-a713-11c5856bfa88` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.direct.position_size.position_size import Protect


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
                style = Protect(position=True, size=True)
                style.apply(rect.component)

                f_style = Protect.from_obj(rect.component)
                assert f_style is not None

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Set Protection of a Shape
^^^^^^^^^^^^^^^^^^^^^^^^^

Setting protection of the shape is done by using the ``Protection`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.position_size.position_size import Protect
        # ... other code

        rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        style = Protect(position=True, size=True)
        style.apply(rect.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape protection can be seen in :numref:`ad9467c1-3b8a-4b49-b626-97564eb8c908`.

.. cssclass:: screen_shot

    .. _ad9467c1-3b8a-4b49-b626-97564eb8c908:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ad9467c1-3b8a-4b49-b626-97564eb8c908
        :alt: Shape with protection set
        :figclass: align-center
        :width: 450px

        Shape with protection set

.. note::

    If the ``position`` is set to ``True`` then the ``size`` will also be set to ``True``.
    Attempting to set ``prop_size`` will be ignored if ``prop_position`` is set to ``True``.

Get Shape Position
^^^^^^^^^^^^^^^^^^

We can get the protection of the shape by using the ``Protect.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.position_size.position_size import Protect
        # ... other code

        # get the position from the shape
        f_style = Protect.from_obj(rect.component)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.position_size.position_size.Protect`
