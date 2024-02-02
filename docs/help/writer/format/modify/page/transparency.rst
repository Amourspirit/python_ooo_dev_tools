.. _help_writer_format_modify_page_transparency:

Write Modify Page Transparency
==============================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.writer.modify.page.transparency.Transparency` and :py:class:`ooodev.format.writer.modify.page.transparency.Gradient` classes are used to modify the Area style values seen in :numref:`235186880-68dd0cdc-8221-40f5-907a-940ec4ef155f` of a Page style.


.. cssclass:: screen_shot

    .. _235186880-68dd0cdc-8221-40f5-907a-940ec4ef155f:
    .. figure:: https://user-images.githubusercontent.com/4193389/235186880-68dd0cdc-8221-40f5-907a-940ec4ef155f.png
        :alt: Writer dialog Transparency default
        :figclass: align-center
        :width: 450px

        Writer dialog Transparency default

Default Page Style Dialog

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        from ooodev.format.writer.modify.page.transparency import (
            Transparency,
            Gradient,
            IntensityRange,
            GradientStyle,
            WriterStylePageKind,
        )
        from ooodev.format.writer.modify.page.area import Color as PageAreaColor
        from ooodev.format import Styler
        from ooodev.utils.color import StandardColor
        from ooodev.office.write import Write
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
           with Lo.Loader(Lo.ConnectPipe()):
                doc = Write.create_doc()
                GUI.set_visible(doc=doc)
                Lo.delay(300)
                GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)

                page_style_kind = WriterStylePageKind.STANDARD
                color_style = PageAreaColor(color=StandardColor.RED, style_name=page_style_kind)
                transparency_style = Transparency(value=85, style_name=page_style_kind)
                Styler.apply(doc, color_style, transparency_style)

                style_obj = Transparency.from_style(doc=doc, style_name=WriterStylePageKind.STANDARD)
                assert style_obj.prop_style_name == str(WriterStylePageKind.STANDARD)

                Lo.delay(1_000)

                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Transparency
------------

The :py:class:`~ooodev.format.writer.modify.page.transparency.Transparency` class is used to modify the transparency of a page style.
The result are seen in :numref:`235190652-995b554d-6db6-443a-a5d4-f8b36de34951`.

Setting Transparency
^^^^^^^^^^^^^^^^^^^^

In this example we will apply a transparency to the page style background color.
The transparency needs to be applied after the page style color as the transparency is applied to the color.
This means the order ``Styler.apply(doc, color_style, transparency_style)`` is important.
The transparency is set to 85% in this example.

.. tabs::

    .. code-tab:: python

        # ... other code

        page_style_kind = WriterStylePageKind.STANDARD
        color_style = PageAreaColor(color=StandardColor.RED, style_name=page_style_kind)
        transparency_style = Transparency(value=85, style_name=page_style_kind)
        Styler.apply(doc, color_style, transparency_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235190652-995b554d-6db6-443a-a5d4-f8b36de34951:
    .. figure:: https://user-images.githubusercontent.com/4193389/235190652-995b554d-6db6-443a-a5d4-f8b36de34951.png
        :alt: Writer dialog Transparency style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Transparency style changed

Getting transparency from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = PageAreaColor.from_style(doc=doc, style_name=page_style_kind)
        assert style_obj.prop_style_name == str(page_style_kind)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Transparency Gradient
---------------------

Setting Transparency Gradient
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.writer.modify.page.transparency.Gradient` class is used to modify the area gradient of a page style.
The result are seen in :numref:`235193804-5d196f94-e80a-4d10-b3f0-625eb7a5880c`.

In this example we will apply a transparency to the page style background color.
The transparency needs to be applied after the page style color as the transparency is applied to the color.
This means the order ``Styler.apply(doc, color_style, para_gradient_style)`` is important.

.. tabs::

    .. code-tab:: python

        # ... other code

        page_style_kind = WriterStylePageKind.STANDARD
        color_style = PageAreaColor(color=StandardColor.GREEN_DARK1, style_name=page_style_kind)
        para_gradient_style = Gradient(
            style=GradientStyle.LINEAR,
            angle=45,
            border=22,
            grad_intensity=IntensityRange(0, 100),
            style_name=page_style_kind,
        )
        Styler.apply(doc, color_style, para_gradient_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _235193804-5d196f94-e80a-4d10-b3f0-625eb7a5880c:
    .. figure:: https://user-images.githubusercontent.com/4193389/235193804-5d196f94-e80a-4d10-b3f0-625eb7a5880c.png
        :alt: Writer dialog Transparency style changed
        :figclass: align-center
        :width: 450px

        Writer dialog Transparency style changed

Getting gradient from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Transparency.from_style(doc=doc, style_name=page_style_kind)
        assert style_obj.prop_style_name == str(page_style_kind)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.writer.modify.page.transparency.Transparency`
        - :py:class:`ooodev.format.writer.modify.page.transparency.Gradient`