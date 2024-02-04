.. _help_draw_format_direct_shape_paragraph_paragraph:

Draw Direct Shape Text Paragraph
================================

Paragraph formatting can be applied by getting the shape text cursor and using the ``append()``, ``append_para()`` or any method that supports setting styles.
This is similar behavior to Writer Direct Paragraph Text. See :ref:`help_writer_format_direct_para` for more information.

Alignment for a paragraph should be set using :ref:`help_draw_format_direct_shape_paragraph_alignment`.

Code Example
------------

.. tabs::

    .. code-tab:: python

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.format.draw.direct.position_size.position_size import Position
        from ooodev.format.writer.direct.char.font import Font, FontEffects
        from ooodev.format.writer.direct.para.indent_space import Indent, Spacing
        from ooodev.format.writer.direct.char.highlight import Highlight
        from ooodev.utils.color import StandardColor
        from ooodev.loader.lo import Lo


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = DrawDoc(Draw.create_draw_doc())
                doc.set_visible()
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_75_PERCENT)

                slide = doc.get_slide()

                width = 100
                height = 80
                x = 0
                y = 0

                rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
                pos = Position(0, 0)
                pos.apply(rect.component)
                cursor = rect.get_shape_text_cursor()
                cursor.append_para(
                    "Hello World!",
                    [
                        Font(b=True, color=StandardColor.GREEN),
                        FontEffects(color=StandardColor.RED),
                        Highlight(color=StandardColor.YELLOW),
                        Indent(first=3.5),
                        Spacing(below=2.5),
                    ],
                )
                cursor.append_para(
                    "Wonderful Day!",
                    [
                        Font(b=False, i=True, color=StandardColor.BLUE_DARK2),
                        Highlight(color=StandardColor.GREEN_LIGHT1),
                        Spacing(below=2.5),
                    ],
                )

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results of running the above code is seen in :numref:`be148d55-3b72-40d0-89d0-b912e739ca19`.

.. cssclass:: screen_shot

    .. _be148d55-3b72-40d0-89d0-b912e739ca19:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/be148d55-3b72-40d0-89d0-b912e739ca19
        :alt: Shape with paragraph and character formatting
        :figclass: align-center
        :width: 450px

        Shape with paragraph and character formatting


.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_direct_para`
        - :ref:`help_writer_format_direct_char`
        - :ref:`help_draw_format_direct_shape_paragraph_alignment`
