.. _help_writer_format_direct_char_font:

Write Direct Character Font Class
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.char.font.Font` class can set the values in :ref:`ss_writer_dialog_char_font`, :ref:`ss_writer_dialog_char_font_effects`
and even more. This class is more of a goto rather than using both :ref:`help_writer_format_direct_char_font_only` and
:ref:`help_writer_format_direct_char_font_effects`.

Setup
-----

.. tabs::

    .. code-tab:: python

        import sys
        from ooodev.format.writer.direct.char.font import Font
        from ooodev.office.write import Write
        from ooodev.utils.color import CommonColor
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo


        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe(Lo.Options(verbose=True))):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.PAGE_WIDTH)
                cursor = Write.get_cursor(doc)
                ft_bold = Font(b=True)
                Write.append(cursor=cursor, text="Have you ", styles=[ft_bold])
                Write.append(
                    cursor=cursor,
                    text="RED",
                    styles=[ft_bold, Font(color=CommonColor.DARK_RED)]
                )
                Write.append_para(
                    cursor=cursor, text=" this?", styles=[Font(b=True, i=True, u=True)]
                )
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

Font bold, italic, underline and color
++++++++++++++++++++++++++++++++++++++

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)
        # font bold instance
        ft_bold = Font(b=True)

        # append bolded text
        Write.append(cursor=cursor, text="Have you ", styles=(ft_bold,))

        # Combine red and bold and two Font instances
        Write.append(
            cursor=cursor, text="RED", styles=(ft_bold, Font(color=CommonColor.DARK_RED))
        )

        # Style text bold, italic,, underline
        Write.append_para(cursor=cursor, text=" this?", styles=(Font(b=True, i=True, u=True),))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Alternatively Font instance can chain together properties.
The last line of code above could have been written.

.. tabs::

    .. code-tab:: python

        Write.append_para(cursor=cursor, text=" this?", styles=(Font().bold.underline.italic,))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

or

.. tabs::

    .. code-tab:: python

        Write.append_para(cursor=cursor, text=" this?", styles=(ft_bold.underline.italic,))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210491001-861ee782-93e2-4836-b508-026697e1437b:
    .. figure:: https://user-images.githubusercontent.com/4193389/210491001-861ee782-93e2-4836-b508-026697e1437b.png
        :alt: Styled Text
        :figclass: align-center

        Styled Text

Font Shadowed
+++++++++++++

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)
        ft = Font(size=17.0, shadowed=True)
        Write.append(cursor=cursor, text="Shadowed", styles=(ft,))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _210492737-ed4cef75-17f3-41ce-9ce7-930320571b32:
    .. figure:: https://user-images.githubusercontent.com/4193389/210492737-ed4cef75-17f3-41ce-9ce7-930320571b32.png
        :alt: Font Shadowed
        :figclass: align-center

        Font Shadowed

Text with hyperlink and superscript
++++++++++++++++++++++++++++++++++++

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.direct.char.hyperlink import Hyperlink, TargetKind
        # ... other code

        cursor = Write.get_cursor(doc)
        ft = Font(color=CommonColor.DARK_GREEN)
        hl = Hyperlink(
            name="machine_learn",
            url="https://en.wikipedia.org//wiki/Machine_learning",
            target=TargetKind.BLANK
        )
        ft_super = Font(name="Liberation Mono", superscript=True)
        Write.append(
            cursor=cursor, text="What do you know about machine learning?", styles=(ft,)
        )
        Write.append(cursor=cursor, text="[", styles=(ft_super,))
        Write.append(cursor=cursor, text="1", styles=(ft_super, hl))
        Write.append_para(cursor=cursor, text="]", styles=(ft_super,))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _211070806-12a3d0a7-6d41-4669-a5d5-955c947a71af:
    .. figure:: https://user-images.githubusercontent.com/4193389/211070806-12a3d0a7-6d41-4669-a5d5-955c947a71af.png
        :alt: What do you know about machine learning?
        :figclass: align-center

        What do you know about machine learning?

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_char_font_only`
        - :ref:`help_writer_format_direct_char_font_effects`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.char.font.Font`