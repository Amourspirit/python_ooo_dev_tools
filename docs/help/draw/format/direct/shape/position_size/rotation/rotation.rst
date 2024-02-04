.. _help_draw_format_direct_shape_position_size_position_rotation:

Draw Direct Shape Position Size - Rotation
==========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.position_size.rotation.Rotation` class is used to modify the values seen in :numref:`2bcb05c7-c393-4f7e-a92b-30c1670069d0` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.format.draw.direct.position_size.rotation import Rotation
        from ooodev.units import Angle100


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
                style = Rotation(rotation=45)
                style.apply(rect.component)

                f_style = Rotation.from_obj(rect.component)
                assert f_style is not None
                assert f_style.prop_rotation == Angle100(4500)

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _2bcb05c7-c393-4f7e-a92b-30c1670069d0:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/2bcb05c7-c393-4f7e-a92b-30c1670069d0
        :alt: Shape Rotation Dialog
        :figclass: align-center
        :width: 450px

        Shape Rotation Dialog

Set Rotation of a Shape
^^^^^^^^^^^^^^^^^^^^^^^

Setting rotation of the shape is done by using the ``Rotation`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.position_size.rotation import Rotation
        # ... other code

        rect = slide.draw_rectangle(
            x=x, y=y, width=width, height=height
        )
        style = Rotation(rotation=45)
        style.apply(rect.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape rotation can be seen in :numref:`cf497479-d821-45d9-850f-dcb9ae8914d8`.

.. cssclass:: screen_shot

    .. _cf497479-d821-45d9-850f-dcb9ae8914d8:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/cf497479-d821-45d9-850f-dcb9ae8914d8
        :alt: Text Shape with Rotation set
        :figclass: align-center

        Text Shape with Rotation set

Get Shape Rotation
^^^^^^^^^^^^^^^^^^

We can get the Adapt properties of the shape by using the ``Rotation.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.position_size.rotation import Rotation
        from ooodev.units import Angle100
        # ... other code

        # get the properties from the shape
        f_style = Rotation.from_obj(rect.component)
        assert f_style.prop_rotation == Angle100(4500)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.position_size.rotation.Rotation`
