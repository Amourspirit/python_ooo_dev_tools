.. _help_writer_format_direct_char_hyperlink:

Write Direct Character Hyperlink Class
======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.direct.char.hyperlink.Hyperlink` class is used to create a hyperlink in a document.

Setting the style
-----------------

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import sys
        from ooodev.format.writer.direct.char.hyperlink import Hyperlink, TargetKind
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo


        def main() -> int:

            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
                cursor = Write.get_cursor(doc)

                ln_name = "OooDev"
                ln_url = "https://python-ooo-dev-tools.readthedocs.io/en/latest/"

                hl = Hyperlink(name=ln_name, url=ln_url)
                Write.append(cursor, "OOO Development Tools", (hl,))
                Write.append_para(cursor, " Docs")

                Write.append(cursor, "Source on Github ")

                pos = Write.get_position(cursor)

                ln_name = "OOODEV_GITHUB"
                ln_url = "https://github.com/Amourspirit/python_ooo_dev_tools"
                hl = Hyperlink(name=ln_name, url=ln_url, target=TargetKind.BLANK)
                Write.append(cursor, "OOO Development Tools on Github", (hl,))

                Lo.delay(2_000)

                # remove the hyperlink
                Write.style_left(cursor=cursor, pos=pos, styles=(hl.empty,))

                # get the hyperlink from the document
                cursor.gotoStart(False)
                cursor.goRight(21, True)
                ooo_dev_hl = Hyperlink.from_obj(cursor)
                cursor.gotoEnd(False)
                print(ooo_dev_hl.prop_url)

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

Adding Hyperlink
++++++++++++++++

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)

        ln_name = "OooDev"
        ln_url = "https://python-ooo-dev-tools.readthedocs.io/en/latest/"
        hl = Hyperlink(name=ln_name, url=ln_url)

        Write.append(cursor, "OOO Development Tools", (hl,))
        Write.append_para(cursor, " Docs")
        # ... other code

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _233806901-da204289-012d-4803-879b-c3a96548f29e:
    .. figure:: https://user-images.githubusercontent.com/4193389/233806901-da204289-012d-4803-879b-c3a96548f29e.png
        :alt: Hyperlink text
        :figclass: align-center

        Hyperlink text

Remove Hyperlink
++++++++++++++++

.. tabs::

    .. code-tab:: python

        cursor = Write.get_cursor(doc)
        # ... other code

        pos = Write.get_position(cursor)

        # ... other code

        Write.style_left(cursor=cursor, pos=pos, styles=(hl.empty,))

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


.. cssclass:: screen_shot

    .. _233806977-2576493c-c24d-4cab-91e8-56be029de9c5:
    .. figure:: https://user-images.githubusercontent.com/4193389/233806977-2576493c-c24d-4cab-91e8-56be029de9c5.png
        :alt: Hyperlink removed
        :figclass: align-center

        Hyperlink removed


Getting the Hyperlink from the text
+++++++++++++++++++++++++++++++++++

.. tabs::

    .. code-tab:: python

        # ... other code
        # get the hyperlink from the document
        cursor.gotoStart(False)
        cursor.goRight(21, True)
        ooo_dev_hl = Hyperlink.from_obj(cursor)
        cursor.gotoEnd(False)
        print(ooo_dev_hl.prop_url)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None



.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.direct.char.hyperlink.Hyperlink`