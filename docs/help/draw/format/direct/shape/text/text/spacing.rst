.. _help_draw_format_direct_shape_text_text_spacing:

Draw Direct Shape Text - Spacing
================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.text.text.Spacing` class is used to modify the values seen in :numref:`80f833c0-0321-4263-a454-b8d2f4ad1dba` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno

        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.format.draw.direct.text.text import Spacing as TextSpacing


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = DrawDoc(Draw.create_draw_doc())
                doc.set_visible()
                Lo.delay(700)
                doc.zoom(ZoomKind.ZOOM_75_PERCENT)

                slide = doc.get_slide()

                width = 50
                height = 50
                x = 10
                y = 10

                rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
                cursor = rect.get_shape_text_cursor()
                cursor.append_para("Hello World!")

                txt_spacing = TextSpacing(left=2.5, right=0.75, top=2.0, bottom=1.7)
                txt_spacing.apply(rect.component)

                f_style = TextSpacing.from_obj(rect.component)
                assert f_style is not None

                # rect.component.TextContourFrame = True  # type: ignore
                # rect.component.TextFitToSize = TextFitToSizeType.PROPORTIONAL  # type: ignore
                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _80f833c0-0321-4263-a454-b8d2f4ad1dba:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/80f833c0-0321-4263-a454-b8d2f4ad1dba
        :alt: Shape Size and Position Dialog
        :figclass: align-center
        :width: 450px

        Shape Size and Position Dialog

Set Spacing of Shape Text
^^^^^^^^^^^^^^^^^^^^^^^^^

Setting spacing of the shape text is done by using the ``Spacing`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.text.text import Spacing as TextSpacing
        # ... other code

        rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        # ... other code
        txt_spacing = TextSpacing(left=2.5, right=0.75, top=2.0, bottom=1.7)
        txt_spacing.apply(rect.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape size can be seen in :numref:`ec7c7c98-bf49-4ff2-9bd8-a7178653b78b`.

.. cssclass:: screen_shot

    .. _ec7c7c98-bf49-4ff2-9bd8-a7178653b78b:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ec7c7c98-bf49-4ff2-9bd8-a7178653b78b
        :alt: Shape with Text Spacing set
        :figclass: align-center
        :width: 450px

        Shape with Text Spacing set

Get Shape Text Spacing
^^^^^^^^^^^^^^^^^^^^^^

We can get the text spacing of the shape by using the ``TextSpacing.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.text.text import Spacing as TextSpacing
        # ... other code

        # get the properties from the shape
        f_style = TextSpacing.from_obj(rect.component)
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.text.text.Spacing`
