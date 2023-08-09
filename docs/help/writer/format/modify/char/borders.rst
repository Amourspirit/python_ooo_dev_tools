.. _help_writer_format_modify_char_borders:

Write Modify Character Borders
==============================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.char.borders.Sides`, :py:class:`ooodev.format.writer.modify.char.borders.Padding`, and :py:class:`ooodev.format.writer.modify.char.borders.Shadow`
classes are used to modify the border values seen in :numref:`234261424-2e08754e-8bb5-4a4a-ab9f-180137f5b50a` of a character border style.


Default Character Borders Style Dialog

.. cssclass:: screen_shot

    .. _234261424-2e08754e-8bb5-4a4a-ab9f-180137f5b50a:
    .. figure:: https://user-images.githubusercontent.com/4193389/234261424-2e08754e-8bb5-4a4a-ab9f-180137f5b50a.png
        :alt: Writer dialog Character Borders default
        :figclass: align-center
        :width: 450px

        Writer dialog Character Borders default


Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.char.borders import Sides, Padding, Shadow
        from ooodev.format.writer.modify.char.borders import Side
        from ooodev.format.writer.modify.char.borders import StyleCharKind
        from ooodev.format.writer.modify.char.borders import BorderLineKind
        from ooodev.format.writer.modify.char.borders import LineSize
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.utils.color import StandardColor

        def main() -> int:
            with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)

                side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED, width=LineSize.MEDIUM)
                sides_style = Sides(border_side=side, style_name=StyleCharKind.EXAMPLE)
                sides_style.apply(doc)

                style_obj = Sides.from_style(doc=doc, style_name=StyleCharKind.EXAMPLE)
                assert style_obj.prop_style_name == str(StyleCharKind.EXAMPLE)
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
        sides_style = Sides(border_side=side, style_name=StyleCharKind.EXAMPLE)
        sides_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234265196-7a12435d-f3f8-4d70-99bb-d2485bf54622:
    .. figure:: https://user-images.githubusercontent.com/4193389/234265196-7a12435d-f3f8-4d70-99bb-d2485bf54622.png
        :alt: Writer dialog Character Borders style sides changed
        :figclass: align-center
        :width: 450px

        Writer dialog Character Borders style sides changed


Getting border sides from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border sides from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Sides.from_style(doc=doc, style_name=StyleCharKind.EXAMPLE)
        assert style_obj.prop_style_name == str(StyleCharKind.EXAMPLE)

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

        padding_style = Padding(left=5, right=5, top=3, bottom=3, style_name=StyleCharKind.EXAMPLE)
        padding_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234267649-bbf10ef2-2b78-4ca8-9c93-5fe4a0248edc:
    .. figure:: https://user-images.githubusercontent.com/4193389/234267649-bbf10ef2-2b78-4ca8-9c93-5fe4a0248edc.png
        :alt: Writer dialog Character Borders style padding changed
        :figclass: align-center
        :width: 450px

        Writer dialog Character Borders style padding changed

Getting border padding from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border padding from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Padding.from_style(doc=doc, style_name=StyleCharKind.EXAMPLE)
        assert style_obj.prop_style_name == str(StyleCharKind.EXAMPLE)

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

        shadow_style = Shadow(color=StandardColor.BLUE_DARK2, width=1.5, style_name=StyleCharKind.EXAMPLE)
        shadow_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _234269140-625bab3f-9e92-444a-a9e2-2f1c16fb1918:
    .. figure:: https://user-images.githubusercontent.com/4193389/234269140-625bab3f-9e92-444a-a9e2-2f1c16fb1918.png
        :alt: Writer dialog Character Borders style shadow changed
        :figclass: align-center
        :width: 450px

        Writer dialog Character Borders style shadow changed

Getting border shadow from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border shadow from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Shadow.from_style(doc=doc, style_name=StyleCharKind.EXAMPLE)
        assert style_obj.prop_style_name == str(StyleCharKind.EXAMPLE)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_writer_format_direct_char_borders`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.char.borders.Padding`
        - :py:class:`ooodev.format.writer.modify.char.borders.Sides`
        - :py:class:`ooodev.format.writer.modify.char.borders.Shadow`