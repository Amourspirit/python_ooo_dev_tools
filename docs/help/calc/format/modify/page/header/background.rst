.. _help_calc_format_modify_page_header_background:

Calc Modify Page Header Background
==================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.calc.modify.page.header.area.Color` and :py:class:`ooodev.format.calc.modify.page.header.area.Img` classes sets the background options for the Calc page.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format import Styler
        from ooodev.format.calc.modify.page.header import Header, CalcStylePageKind
        from ooodev.format.calc.modify.page.header.area import Color as HeaderColor
        from ooodev.format.calc.modify.page.header.area import Img as HeaderImg, PresetImageKind
        from ooodev.office.calc import Calc
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 100)

                header_style = Header(
                    on=True,
                    shared_first=True,
                    shared=True,
                    height=10.0,
                    spacing=3.0,
                    margin_left=1.5,
                    margin_right=2.0,
                    style_name=CalcStylePageKind.DEFAULT,
                )
                header_color_style = HeaderColor(
                    color=StandardColor.GREEN_LIGHT2, style_name=CalcStylePageKind.DEFAULT
                )
                Styler.apply(doc, header_style, header_color_style)

                style_obj = HeaderColor.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
                assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())  

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Header Background Color
-----------------------

Setting the Page Header Background Color
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # .. other code
        header_color_style = HeaderColor(
            color=StandardColor.GREEN_LIGHT2, style_name=CalcStylePageKind.DEFAULT
        )
        Styler.apply(doc, header_style, header_color_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236696549-3d39b26f-4ede-458d-9357-45a54200597c:

    .. figure:: https://user-images.githubusercontent.com/4193389/236696549-3d39b26f-4ede-458d-9357-45a54200597c.png
        :alt: Calc dialog Page Header Background style color modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Header Background style color modified


Getting the background Header color from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # .. other code
        style_obj = HeaderColor.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Header Background Image
-----------------------

Setting the Page Header Background Image
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # .. other code
        header_img_style = HeaderImg.from_preset(
            preset=PresetImageKind.COFFEE_BEANS, style_name=CalcStylePageKind.DEFAULT
        )
        Styler.apply(doc, header_style, header_img_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236696881-a0dd3e2e-b1cd-4640-829f-d2b6983f9552:

    .. figure:: https://user-images.githubusercontent.com/4193389/236696881-a0dd3e2e-b1cd-4640-829f-d2b6983f9552.png
        :alt: Calc dialog Page Header Background style image modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Header Background style image modified

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_modify_page_footer_background`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.calc.modify.page.header.area.Color`
        - :py:class:`ooodev.format.calc.modify.page.header.area.Img`