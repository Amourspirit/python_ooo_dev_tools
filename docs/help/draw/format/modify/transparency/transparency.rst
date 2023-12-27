.. _help_draw_format_modify_transparency_transparency:

Draw Modify Transparency
========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

The :py:class:`ooodev.format.draw.modify.transparency.Transparency`, class is used to modify the values seen in :numref:`c128d890-357d-4c40-afa3-34eec7e69ffd` of a style.

.. cssclass:: screen_shot

    .. _c128d890-357d-4c40-afa3-34eec7e69ffd:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/c128d890-357d-4c40-afa3-34eec7e69ffd
        :alt: Draw dialog Transparency default
        :figclass: align-center
        :width: 450px

        Draw dialog Transparency default


Setting the Transparency
------------------------

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 24, 25, 26, 27, 28, 29

        from __future__ import annotations
        import uno
        from ooodev.draw import Draw, DrawDoc, ZoomKind
        from ooodev.format.draw.modify import FamilyGraphics, DrawStyleFamilyKind
        from ooodev.format.draw.modify.transparency import Transparency, Intensity
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
                style = Transparency(
                    value=Intensity(88),
                    style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
                    style_family=DrawStyleFamilyKind.GRAPHICS,
                )
                doc.apply_styles(style)

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

    .. _377585e8-2815-4044-9763-100b663bdc36:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/377585e8-2815-4044-9763-100b663bdc36
        :alt: Draw dialog Transparency style changed
        :figclass: align-center
        :width: 450px

        Draw dialog Transparency style changed

Shape after applying style.

.. cssclass:: screen_shot

    .. _16c3459a-b219-4739-b903-8ffb21d2c3d7:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/16c3459a-b219-4739-b903-8ffb21d2c3d7
        :alt: Shape after Style applied
        :figclass: align-center

        Shape after Style applied


Getting shadow from a style
---------------------------

.. tabs::

    .. code-tab:: python

        # ... other code

        f_style = Transparency.from_style(
            doc=doc.component,
            style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
            style_family=DrawStyleFamilyKind.GRAPHICS,
        )
        assert f_style.prop_inner.prop_value == Intensity(88)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_draw_format_modify_transparency_gradient`
        - :py:class:`ooodev.format.draw.modify.transparency.Transparency`