.. _help_draw_format_modify_transparency_gradient:

Draw Modify Gradient Transparency
=================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.transparency.Gradient`, class is used to modify the values seen in :numref:`c128d890-357d-4c40-afa3-34eec7e69ffd` of a style.

Setting the Font Effects
------------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 33, 34, 35, 36, 37, 38, 39, 40, 41

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.format.draw.modify import FamilyGraphics, DrawStyleFamilyKind
        from ooodev.format.draw.modify.transparency import Gradient
        from ooodev.format.draw.modify.transparency import GradientStyle, IntensityRange
        from ooodev.format.draw.modify.area.color import Color as AreaColor
        from ooodev.utils.color import StandardColor
        from ooodev.utils.lo import Lo


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = DrawDoc(Draw.create_draw_doc())
                doc.set_visible()
                Lo.delay(700)
                doc.zoom(ZoomKind.ZOOM_75_PERCENT)

                slide = doc.get_slide()

                width = 100
                height = 50
                x = 10
                y = 10

                rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
                rect.set_string("Hello World!")
                style_name_kind = FamilyGraphics.DEFAULT_DRAWING_STYLE
                style_family_kind = DrawStyleFamilyKind.GRAPHICS
                style_color = AreaColor(
                    color=StandardColor.BLUE_LIGHT2, style_name=style_name_kind, style_family=style_family_kind
                )
                style = Gradient(
                    style=GradientStyle.LINEAR,
                    angle=45,
                    border=22,
                    grad_intensity=IntensityRange(0, 100),
                    style_name=style_name_kind,
                    style_family=style_family_kind,
                )
                doc.apply_styles(style_color, style)

                f_style = Gradient.from_style(
                    doc=doc.component,
                    style_name=style_name_kind,
                    style_family=style_family_kind,
                )
                assert f_style.prop_style_name == str(style_name_kind)

                Lo.delay(1_000)
                doc.close_doc()
            return 0


        if __name__ == "__main__":
            raise SystemExit(main())



    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following results in the Draw dialog.

.. cssclass:: screen_shot

    .. _bfada8f6-d48f-40a8-91f0-9a32e1147368:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/bfada8f6-d48f-40a8-91f0-9a32e1147368
        :alt: Draw dialog Gradient Transparency style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Gradient Transparency style changed

Shape after applying style.

.. cssclass:: screen_shot

    .. _d6ecf745-c359-4512-8f4c-fcfa303046f4:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d6ecf745-c359-4512-8f4c-fcfa303046f4
        :alt: Shape after Style applied
        :figclass: align-center

        Shape after Style applied


Getting font effects from a style
---------------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = Gradient.from_style(
            doc=doc.component,
            style_name=style_name_kind,
            style_family=style_family_kind,
        )
        assert f_style.prop_style_name == str(style_name_kind)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_draw_format_modify_transparency_transparency`
        - :ref:`_help_draw_format_modify_area_color`
        - :py:class:`ooodev.format.draw.modify.transparency.Gradient`