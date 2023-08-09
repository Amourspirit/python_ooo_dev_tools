.. _help_writer_format_style_bullet_list:

Write Style BulletList Class
============================

Applying Bullet List Styles can be accomplished using the :py:class:`ooodev.format.writer.style.BulletList` class.

:py:attr:`BulletList.default <ooodev.format.writer.style.BulletList.default>` is used to set the default Bullet List style.

Example Code
------------

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.writer.style import BulletList
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo


        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
                cursor = Write.get_cursor(doc)
                sl = BulletList().list_01
                sl.apply(cursor)
                for i in range(1, 4):
                    Write.append_para(cursor=cursor, text=f"Point {i}")
                BulletList.default.apply(cursor)
                Write.append_para(cursor=cursor, text="Moving On...")

                sl = sl.list_02
                sl.apply(cursor)
                for i in range(1, 4):
                    Write.append_para(cursor=cursor, text=f"Point {i}")
                BulletList.default.apply(cursor)
                Write.append_para(cursor=cursor, text="Moving On...")

                sl = sl.list_03
                sl.apply(cursor)
                for i in range(1, 4):
                    Write.append_para(cursor=cursor, text=f"Point {i}")
                BulletList.default.apply(cursor)
                Write.append_para(cursor=cursor, text="Moving On...")

                sl = sl.num_123
                sl.apply(cursor)
                for i in range(1, 4):
                    Write.append_para(cursor=cursor, text=f"Number Point {i}")
                BulletList.default.apply(cursor)
                Write.append_para(cursor=cursor, text="Moving On...")

                sl = sl.num_IVX
                sl.apply(cursor)
                for i in range(1, 4):
                    Write.append_para(cursor=cursor, text=f"Number Point {i}")
                BulletList.default.apply(cursor)

                Lo.delay(1_000)
                Lo.close_doc(doc)

            return 0


        if __name__ == "__main__":
            SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Output
------

.. cssclass:: screen_shot

    .. _230617401-37f28aab-9d38-497f-8229-92a7b853c384:
    .. figure:: https://user-images.githubusercontent.com/4193389/230617401-37f28aab-9d38-497f-8229-92a7b853c384.png
        :alt: Various Bullet List styles applied to a Writer document
        :figclass: align-center

        Various Bullet List styles applied to a Writer document.

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.office.write.Write`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.style.BulletList`