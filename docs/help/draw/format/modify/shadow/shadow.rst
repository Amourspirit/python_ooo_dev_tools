.. _help_draw_format_modify_shadow_shadow:

Draw Modify Shadow
==================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.shadow.Shadow`, class is used to modify the values seen in :numref:`9a0a1fba-ae02-4f2e-b6e1-eb99299dbf9c` of a style.

.. cssclass:: screen_shot

    .. _9a0a1fba-ae02-4f2e-b6e1-eb99299dbf9c:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9a0a1fba-ae02-4f2e-b6e1-eb99299dbf9c
        :alt: Draw dialog Shadow default
        :figclass: align-center
        :width: 450px

        Draw dialog Shadow default


Setting the Shadow
------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.format.draw.modify import FamilyGraphics, DrawStyleFamilyKind
        from ooodev.format.draw.modify.shadow import Shadow, ShadowLocationKind
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

                style = Shadow(
                    use_shadow=True,
                    location=ShadowLocationKind.BOTTOM_RIGHT,
                    color=StandardColor.YELLOW_LIGHT2,
                    distance=1.5,
                    blur=3,
                    transparency=88,
                    style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
                    style_family=DrawStyleFamilyKind.GRAPHICS,
                )
                doc.apply_styles(style)

                f_style = Shadow.from_style(
                    doc=doc.component, style_name=style.prop_style_name, style_family=style.prop_style_family_name
                )
                assert f_style.prop_style_name == str(FamilyGraphics.DEFAULT_DRAWING_STYLE)
                assert f_style.prop_inner.prop_color == StandardColor.YELLOW_LIGHT2

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

    .. _aa15b057-04c8-4a90-9235-abec6d767f9b:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/aa15b057-04c8-4a90-9235-abec6d767f9b
        :alt: Draw dialog Shadow style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Shadow style changed

Shape after applying style.

.. cssclass:: screen_shot

    .. _8e18313e-9726-43bc-b7b6-5c30007183be:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/8e18313e-9726-43bc-b7b6-5c30007183be
        :alt: Shape after Style applied
        :figclass: align-center

        Shape after Style applied


Getting shadow from a style
---------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = Shadow.from_style(
            doc=doc.component,
            style_name=style.prop_style_name,
            style_family=style.prop_style_family_name
        )
        assert f_style.prop_style_name == str(FamilyGraphics.DEFAULT_DRAWING_STYLE)
        assert f_style.prop_inner.prop_color == StandardColor.YELLOW_LIGHT2

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`ooodev.format.draw.modify.shadow.Shadow`