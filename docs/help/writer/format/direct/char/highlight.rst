.. _help_writer_format_direct_char_highlight:

Write Direct Character Highlight Class
======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.char.highlight.Highlight` is used to set the highlight of one or more characters.

Setting the style
-----------------

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import sys
        from ooodev.format.writer.direct.char.highlight import Highlight
        from ooodev.office.write import Write
        from ooodev.utils.color import CommonColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo


        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.PAGE_WIDTH)
                cursor = Write.get_cursor(doc)
                hl = Highlight(CommonColor.YELLOW_GREEN)
                Write.append(cursor=cursor, text="Highlighting starts ")
                pos = Write.get_position(cursor)
                Write.append(cursor=cursor, text="here", styles=[hl])
                Write.append_para(cursor=cursor, text=".")

                Lo.delay(1_000)

                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            sys.exit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Examples
--------

Highlight text
++++++++++++++

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)
        hl = Highlight(CommonColor.YELLOW_GREEN)
        Write.append(cursor=cursor, text="Highlighting starts ")
        pos = Write.get_position(cursor)
        Write.append(cursor=cursor, text="here", styles=[hl])
        Write.append_para(cursor=cursor, text=".")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _210419171-289189ad-2819-468e-a1c0-ffca1bd478a4:
    .. figure:: https://user-images.githubusercontent.com/4193389/210419171-289189ad-2819-468e-a1c0-ffca1bd478a4.png
        :alt: Highlighting starts here
        :figclass: align-center

        Highlighting starts here

Getting the style from the text.
++++++++++++++++++++++++++++++++

.. tabs::

    .. code-tab:: python

        cursor.gotoStart(False)
        cursor.goRight(pos, False)
        cursor.goRight(4, True)
        hl = Highlight.from_obj(cursor)
        assert hl.prop_color == CommonColor.YELLOW_GREEN
        cursor.gotoEnd(False)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Remove Highlighting
+++++++++++++++++++

.. tabs::

    .. code-tab:: python

        Write.style(pos=pos, length=4, styles=[Highlight().empty])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210423375-1fba1df4-05f4-4195-9a1f-05b6f7acd197:
    .. figure:: https://user-images.githubusercontent.com/4193389/210423375-1fba1df4-05f4-4195-9a1f-05b6f7acd197.png
        :alt: Highlighting starts here no highlight.
        :figclass: align-center

        Highlighting starts here, no highlight.



.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_modify_char_highlight`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.char.highlight.Highlight`