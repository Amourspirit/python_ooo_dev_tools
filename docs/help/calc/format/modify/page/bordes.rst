.. _help_calc_format_modify_page_borders:

Calc Modify Page Borders
========================


.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.calc.modify.page.borders.Sides`, :py:class:`ooodev.format.calc.modify.page.borders.Padding`, and :py:class:`ooodev.format.calc.modify.page.borders.Shadow`
classes are used to modify the border values seen in :numref:`236626523-052afa19-3b16-4bda-a49b-01cda4256afe` of a character border style.


Default Page Borders Style Dialog

.. cssclass:: screen_shot

    .. _236626523-052afa19-3b16-4bda-a49b-01cda4256afe:

    .. figure:: https://user-images.githubusercontent.com/4193389/236626523-052afa19-3b16-4bda-a49b-01cda4256afe.png
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
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.modify.page.borders import Padding, Shadow, Sides
        from ooodev.format.calc.modify.page.borders import BorderLineKind, LineSize
        from ooodev.format.calc.modify.page.borders import Sides, Side, CalcStylePageKind
        from ooodev.utils.color import StandardColor


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 100)

                side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED, width=LineSize.MEDIUM)
                sides_style = Sides(all=side, style_name=CalcStylePageKind.DEFAULT)
                sides_style.apply(doc)

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

Border Sides
------------

Setting Border Sides
^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED, width=LineSize.MEDIUM)
        sides_style = Sides(all=side, style_name=CalcStylePageKind.DEFAULT)
        sides_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236626786-669b2f48-ed2c-4483-8c1e-d370ec16217f:

    .. figure:: https://user-images.githubusercontent.com/4193389/236626786-669b2f48-ed2c-4483-8c1e-d370ec16217f.png
        :alt: Calc dialog Page Style Borders style sides modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Borders style sides modified


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

Border Padding
--------------

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
        padding_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236626903-4335f208-fb12-4a48-b0b3-fa39c2e06f17:

    .. figure:: https://user-images.githubusercontent.com/4193389/236626903-4335f208-fb12-4a48-b0b3-fa39c2e06f17.png
        :alt: Calc dialog Page Style Borders style padding modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Borders style padding modified

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

Border Shadow
-------------

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
        shadow_style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236627071-885795c6-4fff-4574-8244-6702486e949e:

    .. figure:: https://user-images.githubusercontent.com/4193389/236627071-885795c6-4fff-4574-8244-6702486e949e.png
        :alt: Calc dialog Page Style Borders style shadow modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Borders style shadow modified

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
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.calc.modify.page.borders.Padding`
        - :py:class:`ooodev.format.calc.modify.page.borders.Sides`
        - :py:class:`ooodev.format.calc.modify.page.borders.Shadow`