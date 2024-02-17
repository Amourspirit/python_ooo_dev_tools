.. _help_calc_format_modify_page_header_borders:

Calc Modify Page Header Borders
===============================


.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.calc.modify.page.header.borders.Sides`, :py:class:`ooodev.format.calc.modify.page.header.borders.Padding`, and :py:class:`ooodev.format.calc.modify.page.header.borders.Shadow`
classes are used to modify the border values seen in :numref:`236699240-c0a869c1-67aa-4b14-94e3-7a4dad48a7d5` of a character border style.


Default Page Header Borders Style Dialog

.. cssclass:: screen_shot

    .. _236699240-c0a869c1-67aa-4b14-94e3-7a4dad48a7d5:

    .. figure:: https://user-images.githubusercontent.com/4193389/236699240-c0a869c1-67aa-4b14-94e3-7a4dad48a7d5.png
        :alt: Calc dialog Page Style Borders default
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Borders default


Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format import Styler
        from ooodev.format.calc.modify.page.header import Header, CalcStylePageKind
        from ooodev.format.calc.modify.page.header.borders import BorderLineKind, LineSize
        from ooodev.format.calc.modify.page.header.borders import Padding, Shadow, Sides
        from ooodev.format.calc.modify.page.header.borders import Sides, Side
        from ooodev.office.calc import Calc
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo


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
                side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED, width=LineSize.MEDIUM)
                header_sides_style = Sides(all=side, style_name=CalcStylePageKind.DEFAULT)
                Styler.apply(doc, header_style, header_sides_style)

                style_obj = Sides.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
                assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Header Border Sides
-------------------

Setting Border Sides
^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED, width=LineSize.MEDIUM)
        header_sides_style = Sides(all=side, style_name=CalcStylePageKind.DEFAULT)
        Styler.apply(doc, header_style, header_sides_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236699440-ac37907b-d693-4f8b-a223-c62749f8a609:

    .. figure:: https://user-images.githubusercontent.com/4193389/236699440-ac37907b-d693-4f8b-a223-c62749f8a609.png
        :alt: Calc dialog Page Style Header Borders style sides modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Header Borders style sides modified


Getting border sides from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border sides from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Sides.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Header Border Padding
---------------------

Setting Border Padding
^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        padding_style = Padding(
            left=5,
            right=5,
            top=3,
            bottom=3,
            style_name=CalcStylePageKind.DEFAULT,
        )
        Styler.apply(doc, header_style, padding_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236699612-cdeab377-1081-4308-9aee-7076b7a99817:

    .. figure:: https://user-images.githubusercontent.com/4193389/236699612-cdeab377-1081-4308-9aee-7076b7a99817.png
        :alt: Calc dialog Page Header Style Borders style padding modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Header Style Borders style padding modified

Getting border padding from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border padding from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Padding.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Header Border Shadow
--------------------

Setting Border Shadow
^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        shadow_style = Shadow(
            color=StandardColor.BLUE_DARK2,
            width=1.5,
            style_name=CalcStylePageKind.DEFAULT,
        )
        Styler.apply(doc, header_style, shadow_style)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236699766-e4cdd9ab-0e51-4a88-a0b6-30870862d076:

    .. figure:: https://user-images.githubusercontent.com/4193389/236699766-e4cdd9ab-0e51-4a88-a0b6-30870862d076.png
        :alt: Calc dialog Page Header Style Borders style shadow modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Header Style Borders style shadow modified

Getting border shadow from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

We can get the border shadow from the document.

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Shadow.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_calc_format_modify_cell_borders`
        - :ref:`help_calc_format_modify_page_footer_borders`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.calc.modify.page.header.borders.Padding`
        - :py:class:`ooodev.format.calc.modify.page.header.borders.Sides`
        - :py:class:`ooodev.format.calc.modify.page.header.borders.Shadow`