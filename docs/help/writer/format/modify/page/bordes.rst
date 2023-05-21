.. _help_writer_format_modify_page_borders:

Write Modify Page Borders
=========================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.page.borders.Sides`, :py:class:`ooodev.format.writer.modify.page.borders.Padding`, and :py:class:`ooodev.format.writer.modify.page.borders.Shadow`
classes are used to modify the border values seen in :numref:`235786832-a11f7b0b-4f93-4af4-8839-b1147bb4ecd4` of a character border style.


Default Page Borders Style Dialog

.. cssclass:: screen_shot

    .. _235786832-a11f7b0b-4f93-4af4-8839-b1147bb4ecd4:
    .. figure:: https://user-images.githubusercontent.com/4193389/235786832-a11f7b0b-4f93-4af4-8839-b1147bb4ecd4.png
        :alt: Writer dialog Page Header Borders default
        :figclass: align-center
        :width: 450px

        Writer dialog Page Header Borders default


Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.borders import Padding, Shadow, Sides, WriterStylePageKind
        from ooodev.format.writer.modify.page.borders import BorderLineKind, LineSize, Side
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.utils.color import StandardColor

        def main() -> int:
          with Lo.Loader(Lo.ConnectPipe()):
            doc = Write.create_doc()
            GUI.set_visible(doc)
            Lo.delay(300)
            GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

            side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED, width=LineSize.MEDIUM)
            sides_style = Sides(all=side, style_name=WriterStylePageKind.STANDARD)
            sides_style.apply(doc)

            style_obj = Sides.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
            assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)
            Lo.delay(1_000)

            Lo.close_doc(doc)
        return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Border Sides
------------

Setting Border Sides
^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED, width=LineSize.MEDIUM)
        sides_style = Sides(all=side, style_name=WriterStylePageKind.STANDARD)
        sides_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235788486-b4414e96-c704-4a55-9c46-4fdb6f4d8a96:
    .. figure:: https://user-images.githubusercontent.com/4193389/235788486-b4414e96-c704-4a55-9c46-4fdb6f4d8a96.png
        :alt: Writer dialog Page Borders style sides changed
        :figclass: align-center
        :width: 450px

        Writer dialog Page Borders style sides changed


Getting border sides from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Sides.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Border Padding
--------------

Setting Border Padding
^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        padding_style = Padding(
            left=5, right=5, top=3, bottom=3, style_name=WriterStylePageKind.STANDARD
        )
        padding_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235789495-b5c1f937-6140-45f4-a528-3c2c6b98114d:
    .. figure:: https://user-images.githubusercontent.com/4193389/235789495-b5c1f937-6140-45f4-a528-3c2c6b98114d.png
        :alt: Writer dialog Page Borders style padding changed
        :figclass: align-center
        :width: 450px

        Writer dialog Page Borders style padding changed

Getting border padding from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border padding from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Padding.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Border Shadow
-------------

Setting Border Shadow
^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        shadow_style = Shadow(
            color=StandardColor.BLUE_DARK2, width=1.5, style_name=WriterStylePageKind.STANDARD
        )
        shadow_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235790113-919dce2b-fec3-4513-9e17-028e7cb7092e:
    .. figure:: https://user-images.githubusercontent.com/4193389/235790113-919dce2b-fec3-4513-9e17-028e7cb7092e.png
        :alt: Writer dialog Page Borders style shadow changed
        :figclass: align-center
        :width: 450px

        Writer dialog Page Borders style shadow changed

Getting border shadow from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border shadow from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Shadow.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
        assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_para_borders`
        - :ref:`help_writer_format_modify_para_borders`
        - :ref:`help_writer_format_modify_page_header_borders`
        - :ref:`help_writer_format_modify_page_footer_borders`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.page.borders.Padding`
        - :py:class:`ooodev.format.writer.modify.page.borders.Sides`
        - :py:class:`ooodev.format.writer.modify.page.borders.Shadow`