.. _help_writer_format_direct_para_area_color:

Write Direct Paragraph Area Color Class
=======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.writer.direct.para.area.Color` class is used to set the paragraph background area color.

Setting the style
-----------------

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import sys
        from ooodev.office.write import Write
        from ooodev.utils.color import CommonColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.writer.direct.para.area import Color as ParaBgColor


        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
                cursor = Write.get_cursor(doc)
                fc = ParaBgColor(CommonColor.YELLOW_GREEN)
                Write.append_para(cursor=cursor, text="Fill Color starts Here", styles=[fc])

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

Fill Color a single Paragraph
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)
        fc = ParaBgColor(CommonColor.YELLOW_GREEN)
        Write.append_para(cursor=cursor, text="Fill Color starts Here", styles=[fc])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _212849953-4d194f87-080e-4417-b376-41b5b68ef744:
    .. figure:: https://user-images.githubusercontent.com/4193389/212849953-4d194f87-080e-4417-b376-41b5b68ef744.png
        :alt: Paragraph with background color
        :figclass: align-center

        Paragraph with background color.

.. cssclass:: screen_shot

    .. _212850105-ab0afde4-ff6f-42e1-ac54-78c0ad4cae04:
    .. figure:: https://user-images.githubusercontent.com/4193389/212850105-ab0afde4-ff6f-42e1-ac54-78c0ad4cae04.png
        :alt: Paragraph area color dialog
        :figclass: align-center

        Paragraph area color dialog.

Fill Color Multiple paragraphs
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)
        fc = ParaBgColor(CommonColor.YELLOW_GREEN)
        Write.append_para(cursor=cursor, text="Fill Color starts Here", styles=[fc])
        fc = ParaBgColor(CommonColor.LIGHT_BLUE)
        Write.append_para(cursor=cursor, text="And today Ends Here", styles=[fc])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _212850641-79f87983-987e-404f-bccc-3d1740d8e361:
    .. figure:: https://user-images.githubusercontent.com/4193389/212850641-79f87983-987e-404f-bccc-3d1740d8e361.png
        :alt: Paragraph with background color
        :figclass: align-center

        Paragraph with background color.


Apply Fill cursor to Cursor
^^^^^^^^^^^^^^^^^^^^^^^^^^^

A Fill Color can be set on the cursor and then it remains until it is removed.

The fill color can be cleared by using :py:attr:`ParaStyle.default <ooodev.format.writer.style.Para.default>` values.


.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.style.para import Para as ParaStyle
        # ... other code

        cursor = Write.get_cursor(doc)
        fc = ParaBgColor(CommonColor.YELLOW_GREEN)
        fc.apply(cursor.TextParagraph)
        Write.append_para(cursor=cursor, text="Fill Color starts Here")
        Write.append_para(cursor=cursor, text="And today Ends Here")
        ParaStyle.default.apply(cursor)
        Write.append_para(cursor=cursor, text="Nothing to report")

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

.. cssclass:: screen_shot

    .. _212851968-ab25ac9b-04a0-40aa-b3f5-808f2aa492f9:
    .. figure:: https://user-images.githubusercontent.com/4193389/212851968-ab25ac9b-04a0-40aa-b3f5-808f2aa492f9.png
        :alt: Paragraph style reset
        :figclass: align-center

        Paragraph style reset.

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_writer_format_style_para_reset_default`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.para.area.Color`