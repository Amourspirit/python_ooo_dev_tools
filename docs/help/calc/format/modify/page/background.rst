.. _help_calc_format_modify_page_background:

Calc Modify Page Background
===========================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.calc.modify.page.area.Color` and :py:class:`ooodev.format.calc.modify.page.area.Img` classes sets the fill options for the Calc page.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.calc.modify.page.area import Color as PageStyleColor, CalcStylePageKind
        from ooodev.format.calc.modify.page.area import Img as PageStyleImg, PresetImageKind
        from ooodev.utils.color import StandardColor

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 100)

                page_color_style = PageStyleColor(
                    color=StandardColor.GREEN_LIGHT2, style_name=CalcStylePageKind.DEFAULT
                )
                page_color_style.apply(doc)

                style_obj = PageStyleColor.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
                assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Background Color
----------------

Setting the Page Background Color
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # .. other code
        page_color_style = PageStyleColor(
            color=StandardColor.GREEN_LIGHT2, style_name=CalcStylePageKind.DEFAULT
        )
        page_color_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236639347-f8ea096c-7f23-4d0c-a1f5-96d997c4727f:

    .. figure:: https://user-images.githubusercontent.com/4193389/236639347-f8ea096c-7f23-4d0c-a1f5-96d997c4727f.png
        :alt: Calc dialog Page Style Background style color modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Background style color modified


Getting the background color from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # .. other code
        style_obj = PageStyleColor.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Background Image
----------------

Setting the Page Background Image
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # .. other code
        page_img_style = PageStyleImg.from_preset(
            preset=PresetImageKind.COFFEE_BEANS, style_name=CalcStylePageKind.DEFAULT
        )
        page_img_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236640290-799efe37-4239-48e2-ab6a-8f6aed99f7c2:

    .. figure:: https://user-images.githubusercontent.com/4193389/236640290-799efe37-4239-48e2-ab6a-8f6aed99f7c2.png
        :alt: Calc dialog Page Style Background style image modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Background style image modified

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.calc.modify.page.area.Color`
        - :py:class:`ooodev.format.calc.modify.page.area.Img`