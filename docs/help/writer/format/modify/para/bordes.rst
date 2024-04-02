.. _help_writer_format_modify_para_borders:

Write Modify Paragraph Borders
==============================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.para.borders.Sides`, :py:class:`ooodev.format.writer.modify.char.borders.Padding`, and :py:class:`ooodev.format.writer.modify.para.borders.Shadow`
classes are used to modify the border values seen in :numref:`234408779-4e2e1ee9-6582-41df-b702-f4353ca8d8e2` of a character border style.


Default Paragraph Borders Style Dialog

.. cssclass:: screen_shot

    .. _234408779-4e2e1ee9-6582-41df-b702-f4353ca8d8e2:
    .. figure:: https://user-images.githubusercontent.com/4193389/234408779-4e2e1ee9-6582-41df-b702-f4353ca8d8e2.png
        :alt: Writer dialog Paragraph Borders default
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Borders default


Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.writer.modify.para.borders import Padding, Shadow, Sides
        from ooodev.format.writer.modify.para.borders import BorderLineKind, LineSize
        from ooodev.format.writer.modify.para.borders import StyleParaKind, Side
        from ooodev.office.write import Write
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED, width=LineSize.MEDIUM)
                sides_style = Sides(all=side, style_name=StyleParaKind.STANDARD)
                sides_style.apply(doc)

                style_obj = Sides.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
                assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)
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
        sides_style = Sides(all=side, style_name=StyleParaKind.STANDARD)
        sides_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234409135-3e1cd6d5-f1e9-4d2f-bb86-b51bdf1fb486:

    .. figure:: https://user-images.githubusercontent.com/4193389/234409135-3e1cd6d5-f1e9-4d2f-bb86-b51bdf1fb486.png
        :alt: Writer dialog Paragraph Borders style sides modified
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Borders style sides modified


Getting border sides from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border sides from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Sides.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

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

        padding_style = Padding(left=5, right=5, top=3, bottom=3, style_name=StyleParaKind.STANDARD)
        padding_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234410090-e24a79d7-c2f5-460b-b229-02daf243710f:

    .. figure:: https://user-images.githubusercontent.com/4193389/234410090-e24a79d7-c2f5-460b-b229-02daf243710f.png
        :alt: Writer dialog Paragraph Borders style padding modified
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Borders style padding modified

Getting border padding from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border padding from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Padding.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

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

        shadow_style = Shadow(color=StandardColor.BLUE_DARK2, width=1.5, style_name=StyleParaKind.STANDARD)
        shadow_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234410957-55eedfcc-9032-48b1-a660-7dffa5eb5d8f:

    .. figure:: https://user-images.githubusercontent.com/4193389/234410957-55eedfcc-9032-48b1-a660-7dffa5eb5d8f.png
        :alt: Writer dialog Paragraph Borders style shadow modified
        :figclass: align-center
        :width: 450px

        Writer dialog Paragraph Borders style shadow modified

Getting border shadow from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border shadow from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Shadow.from_style(doc=doc, style_name=StyleParaKind.STANDARD)
        assert style_obj.prop_style_name == str(StyleParaKind.STANDARD)

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
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.writer.modify.para.borders.Padding`
        - :py:class:`ooodev.format.writer.modify.para.borders.Sides`
        - :py:class:`ooodev.format.writer.modify.para.borders.Shadow`