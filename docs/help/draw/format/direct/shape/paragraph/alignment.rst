.. _help_draw_format_direct_shape_paragraph_alignment:

Draw Direct Shape Text Paragraph Alignment
==========================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.para.alignment.Alignment` class is used to modify the values seen in :numref:`87d1a548-322e-4ad7-bc35-1397d6617d1a` of a shape.

Setup
-----

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.format.draw.direct.text import TextColumns
        from ooodev.format.draw.direct.text.text import TextAnchor, ShapeBasePointKind
        from ooodev.format.draw.direct.para.alignment import Alignment, ParagraphAdjust
        from ooodev.format import Styler


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = DrawDoc(Draw.create_draw_doc())
                doc.set_visible()
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_75_PERCENT)

                slide = doc.get_slide()

                width = 100
                height = 50
                x = 10
                y = 10

                rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
                sb = []
                for _ in range(12):
                    sb.append("Hello World!")
                rect.set_string("\n".join(sb))

                anchor = TextAnchor(anchor_point=ShapeBasePointKind.CENTER, full_width=True)
                align = Alignment(align=ParagraphAdjust.CENTER)
                txt_cols = TextColumns(col_count=2, spacing=0.5)
                Styler.apply(rect.component, anchor, align, txt_cols)

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _87d1a548-322e-4ad7-bc35-1397d6617d1a:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/87d1a548-322e-4ad7-bc35-1397d6617d1a
        :alt: Shape Text Columns Dialog
        :figclass: align-center
        :width: 450px

        Shape Text Columns Dialog

Set shapes text alignment
^^^^^^^^^^^^^^^^^^^^^^^^^

Setting Text Columns of the shape text is done by using the ``Alignment`` class.

The ``Styler`` class is used to apply several styles to the shape at one time.

.. tabs::

    .. code-tab:: python

        rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        # ... other code
        anchor = TextAnchor(anchor_point=ShapeBasePointKind.CENTER, full_width=True)
        align = Alignment(align=ParagraphAdjust.CENTER)
        txt_cols = TextColumns(col_count=2, spacing=0.5)
        Styler.apply(rect.component, anchor, align, txt_cols)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The dialog results of the setting the shape text columns can be seen in :numref:`87d1a548-322e-4ad7-bc35-1397d6617d1a`.

Get Shape Text Alignment
^^^^^^^^^^^^^^^^^^^^^^^^

We can get the text alignment of the shape by using the ``Alignment.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        # get the properties from the shape
        f_style = Alignment.from_obj(rect.component)
        assert f_style.prop_align == ParagraphAdjust.CENTER

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.para.alignment.Alignment`
        - :ref:`help_draw_format_direct_shape_text_text_columns`
        - :ref:`help_draw_format_direct_shape_text_text_anchor`
        - :ref:`help_draw_format_direct_shape_paragraph_paragraph`
