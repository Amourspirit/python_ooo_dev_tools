.. _help_calc_format_modify_page_footer_background:

Calc Modify Page Footer Background
==================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.calc.modify.page.footer.area.Color` and :py:class:`ooodev.format.calc.modify.page.footer.area.Img` classes sets the background options for the Calc page.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format import Styler
        from ooodev.format.calc.modify.page.footer import Footer, CalcStylePageKind
        from ooodev.format.calc.modify.page.footer.area import Color as FooterColor
        from ooodev.format.calc.modify.page.footer.area import Img as FooterImg, PresetImageKind
        from ooodev.office.calc import Calc
        from ooodev.utils.color import StandardColor
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 100)

                footer_style = Footer(
                    on=True,
                    shared_first=True,
                    shared=True,
                    height=10.0,
                    spacing=3.0,
                    margin_left=1.5,
                    margin_right=2.0,
                    style_name=CalcStylePageKind.DEFAULT,
                )
                footer_color_style = FooterColor(
                    color=StandardColor.GREEN_LIGHT2, style_name=CalcStylePageKind.DEFAULT
                )
                Styler.apply(doc, footer_style, footer_color_style)

                style_obj = FooterColor.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
                assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())  

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Footer Background Color
-----------------------

Setting the Page Footer Background Color
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # .. other code
        footer_color_style = FooterColor(
            color=StandardColor.GREEN_LIGHT2, style_name=CalcStylePageKind.DEFAULT
        )
        Styler.apply(doc, footer_style, footer_color_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236696549-3d39b26f-4ede-458d-9357-45a54200597c_2:

    .. figure:: https://user-images.githubusercontent.com/4193389/236696549-3d39b26f-4ede-458d-9357-45a54200597c.png
        :alt: Calc dialog Page Footer Background style color modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Footer Background style color modified


Getting the background Footer color from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # .. other code
        style_obj = FooterColor.from_style(
            doc=doc, style_name=CalcStylePageKind.DEFAULT
        )
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Footer Background Image
-----------------------

Setting the Page Footer Background Image
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # .. other code
        footer_img_style = FooterImg.from_preset(
            preset=PresetImageKind.COFFEE_BEANS, style_name=CalcStylePageKind.DEFAULT
        )
        Styler.apply(doc, footer_style, footer_img_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236696881-a0dd3e2e-b1cd-4640-829f-d2b6983f9552_2:

    .. figure:: https://user-images.githubusercontent.com/4193389/236696881-a0dd3e2e-b1cd-4640-829f-d2b6983f9552.png
        :alt: Calc dialog Page Footer Background style image modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Footer Background style image modified

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_modify_page_header_background`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.calc.modify.page.footer.area.Color`
        - :py:class:`ooodev.format.calc.modify.page.footer.area.Img`