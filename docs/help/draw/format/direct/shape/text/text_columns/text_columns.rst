.. _help_draw_format_direct_shape_text_text_columns:

Draw Direct Shape Text - Columns
================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.direct.text.TextColumns` class is used to modify the values seen in :numref:`ea668881-f18d-4a71-b046-82d4900ec075` of a shape.

For an example of using text columns with Writer see `Live LibreOffice Python UNO - Text Columns Example <https://github.com/Amourspirit/python-ooouno-ex/tree/main/ex/auto/writer/odev_text_columns>`__.

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

    .. _ea668881-f18d-4a71-b046-82d4900ec075:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ea668881-f18d-4a71-b046-82d4900ec075
        :alt: Shape Text Columns Dialog
        :figclass: align-center

        Shape Text Columns Dialog

Set Text Columns of Shape Text
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Setting Text Columns of the shape text is done by using the ``TextColumns`` class.

Note that the anchor and alignment must also be to see the results of the text columns.
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

The dialog results of the setting the shape text columns can be seen in :numref:`40d2cce7-5616-49f7-8e82-8332585b1f15`.

.. cssclass:: screen_shot

    .. _40d2cce7-5616-49f7-8e82-8332585b1f15:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/40d2cce7-5616-49f7-8e82-8332585b1f15
        :alt: Shape Text Columns Dialog
        :figclass: align-center

        Shape with Text Anchor Set

The output results of the setting the shape text columns can be seen in :numref:`4d7ed1ec-1e04-4d0b-8de6-52cff6688997`.

.. cssclass:: screen_shot

    .. _4d7ed1ec-1e04-4d0b-8de6-52cff6688997:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4d7ed1ec-1e04-4d0b-8de6-52cff6688997
        :alt: Shape Text Columns Screen Shot
        :figclass: align-center
        :width: 450px

        Shape Text Columns Screen Shot

Get Shape Text Columns
^^^^^^^^^^^^^^^^^^^^^^

We can get the text columns of the shape by using the ``TextColumns.from_obj()`` method.

.. tabs::

    .. code-tab:: python

        # get the properties from the shape
        f_style = TextColumns.from_obj(rect.component)
        assert f_style.prop_col_count == 2

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`ooodev.format.draw.direct.text.TextColumns`
        - :ref:`help_draw_format_direct_shape_text_text_anchor`
        - :ref:`help_draw_format_direct_shape_paragraph_alignment`
