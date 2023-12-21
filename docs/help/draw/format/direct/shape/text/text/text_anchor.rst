.. _help_draw_format_direct_shape_text_text_anchor:

Draw Direct Shape Text - Anchor
===============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.text.text.TextAnchor` class is used to modify the values seen in :numref:`20b21690-fd62-4f40-b14f-f8c936983d42` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno

        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.utils.lo import Lo
        from ooodev.format.draw.direct.text.text import Spacing as TextSpacing
        from ooodev.format.draw.direct.text.text import TextAnchor, ShapeBasePointKind


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = DrawDoc(Draw.create_draw_doc())
                doc.set_visible()
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_75_PERCENT)

                slide = doc.get_slide()

                width = 50
                height = 50
                x = 10
                y = 10

                rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
                cursor = rect.get_shape_text_cursor()
                cursor.append_para("Hello World!")

                txt_anchor = TextAnchor(
                    anchor_point=ShapeBasePointKind.TOP_CENTER, full_width=True
                )
                txt_anchor.apply(rect.component)

                f_style = TextAnchor.from_obj(rect.component)
                assert f_style.prop_full_width is True
                assert f_style.prop_anchor_point == ShapeBasePointKind.TOP_CENTER

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _20b21690-fd62-4f40-b14f-f8c936983d42:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/20b21690-fd62-4f40-b14f-f8c936983d42
        :alt: Shape Text Anchor Point Dialog
        :figclass: align-center
        :width: 450px

        Shape Text Anchor Point Dialog

Set Text Anchor of Shape Text
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Setting Text Anchor Point of the shape text is done by using the ``TextAnchor`` class.

.. tabs::

    .. code-tab:: python

        
        from ooodev.format.draw.direct.text.text import TextAnchor, ShapeBasePointKind
        # ... other code

        rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        # ... other code
        txt_anchor = TextAnchor(
            anchor_point=ShapeBasePointKind.TOP_CENTER, full_width=True
        )
        txt_anchor.apply(rect.component)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of the setting the shape text anchor can be seen in :numref:`9e4f4ea9-cab7-4774-a835-163e6f15144a`.

.. cssclass:: screen_shot

    .. _9e4f4ea9-cab7-4774-a835-163e6f15144a:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9e4f4ea9-cab7-4774-a835-163e6f15144a
        :alt: Shape with Text Anchor Set
        :figclass: align-center
        :width: 450px

        Shape with Text Anchor Set

Get Shape Text Anchor
^^^^^^^^^^^^^^^^^^^^^

We can get the text anchor of the shape by using the ``TextAnchor.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        from ooodev.format.draw.direct.text.text import TextAnchor
        # ... other code

        # get the properties from the shape
        f_style = TextAnchor.from_obj(rect.component)
        assert f_style.prop_full_width is True
        assert f_style.prop_anchor_point == ShapeBasePointKind.TOP_CENTER

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.text.text.TextAnchor`
