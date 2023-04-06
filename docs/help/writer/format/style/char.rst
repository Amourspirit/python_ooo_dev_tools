.. _help_writer_format_style_char:

Write Style Char Class
======================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2


Applying Character Styles can be accomplished using the :py:class:`ooodev.format.writer.style.Char` class.

Setup
-----

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.style import Char as StyleChar, StyleCharKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo


        def main() -> int:

            with Lo.Loader(Lo.ConnectSocket()):
                doc = Write.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                GUI.zoom(GUI.ZoomEnum.ZOOM_100_PERCENT)

                cursor = Write.get_cursor(doc)
                Write.append(
                    cursor=cursor,
                    text="The quick fox jumped over the lazy dog.",
                    styles=[StyleChar().quotation],
                )
                cursor.goLeft(25, False)
                cursor.goRight(6, True)
                StyleChar(name=StyleCharKind.STRONG_EMPHASIS).apply(cursor)
                cursor.gotoEnd(False)
                Write.end_paragraph(cursor)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply Style
-----------

.. tabs::

    .. code-tab:: python

        # ... other code
        cursor = Write.get_cursor(doc)
        Write.append(
            cursor=cursor,
            text="The quick fox jumped over the lazy dog.",
            styles=[StyleChar().quotation],
        )
        # select the word "jumped"
        cursor.goLeft(25, False)
        cursor.goRight(6, True)
        StyleChar(name=StyleCharKind.STRONG_EMPHASIS).apply(cursor)
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)
        # ... other code

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _214347692-e5812d4b-f6d4-40a6-bcfe-e5d33be8772a:
    .. figure:: https://user-images.githubusercontent.com/4193389/214347692-e5812d4b-f6d4-40a6-bcfe-e5d33be8772a.png
        :alt: Sentence with style char
        :figclass: align-center

        Sentence with style char.

Get Style from Cursor
---------------------

.. tabs::

    .. code-tab:: python

        # ... other code
        # select the word "jumped"
        cursor.gotoStart(False)
        cursor.goRight(14, False)
        cursor.goRight(6, True)
        style = StyleChar.from_obj(cursor)
        cursor.gotoEnd(False)
        assert style.prop_name == "Strong Emphasis"
        # ... other code

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. seealso::

    .. cssclass:: ul-list

        - :py:class:`~ooodev.office.write.Write`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
