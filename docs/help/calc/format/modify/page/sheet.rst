.. _help_calc_format_modify_page_sheet:

Calc Modify Page Sheet
======================


.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 3

Overview
--------

The :py:class:`ooodev.format.calc.modify.page.sheet.Order` and :py:class:`ooodev.format.calc.modify.page.sheet.Printing` classes sets the sheet order and print options for the Calc page.

The Scale section is set using the :py:class:`ooodev.format.calc.modify.page.sheet.ScaleReduceEnlarge`, :py:class:`ooodev.format.calc.modify.page.sheet.ScaleNumOfPages` and :py:class:`ooodev.format.calc.modify.page.sheet.ScalePagesWidthHeight` classes.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.loader.lo import Lo
        from ooodev.format.calc.modify.page.sheet import Order, Printing
        from ooodev.format.calc.modify.page.sheet import CalcStylePageKind

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 100)

                style = Order(top_btm=False, first_pg=0, style_name=CalcStylePageKind.DEFAULT)
                style.apply(doc)

                style_obj = Order.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
                assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Page Order
----------

The :py:class:`~ooodev.format.calc.modify.page.sheet.Order` class sets the page order of the page sheet style.

Setting the Page Order
^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # .. other code
        style = Order(top_btm=False, first_pg=0, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236641402-fbcc9fc9-1438-465e-bdfe-2bf0b9fa4a0b:

    .. figure:: https://user-images.githubusercontent.com/4193389/236641402-fbcc9fc9-1438-465e-bdfe-2bf0b9fa4a0b.png
        :alt: Calc dialog Page Style Sheet Order modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Sheet Order modified


Getting the page order from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # .. other code
        style_obj = Order.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Print Options
-------------

The :py:class:`~ooodev.format.calc.modify.page.sheet.Printing` class sets the print options of the page sheet style.


Setting the Page Print Options
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.modify.page.sheet import Printing, CalcStylePageKind
    
        # .. other code
        style = Printing(
            header=False,
            grid=False,
            chart=False,
            drawing=False,
            style_name=CalcStylePageKind.DEFAULT,
        )
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236646444-0e36e8cb-f3b6-4699-8357-da4bdec6b748:

    .. figure:: https://user-images.githubusercontent.com/4193389/236646444-0e36e8cb-f3b6-4699-8357-da4bdec6b748.png
        :alt: Calc dialog Page Style Sheet Print Options style modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Sheet Print Options style modified

Getting the print options from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # .. other code
        style_obj = Printing.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Scale Options
-------------

Reduce/Enlarge
^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.calc.modify.page.sheet.ScaleReduceEnlarge` class sets the scale reduce/enlarge settings of the page sheet style.

Setting the Page Scale Reduce/Enlarge
"""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.modify.page.sheet import ScaleReduceEnlarge, CalcStylePageKind
    
        # .. other code
        style = ScaleReduceEnlarge(factor=200, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236646611-ee6a4036-7655-4d34-b56f-60deb2074dc3:

    .. figure:: https://user-images.githubusercontent.com/4193389/236646611-ee6a4036-7655-4d34-b56f-60deb2074dc3.png
        :alt: Calc dialog Page Style Sheet Scale style image modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Sheet Scale style image modified

Getting the page scale reduce/enlarge from a style
""""""""""""""""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # .. other code
        style_obj = ScaleReduceEnlarge.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Width/Height
^^^^^^^^^^^^

The :py:class:`~ooodev.format.calc.modify.page.sheet.ScalePagesWidthHeight` class sets the scale width/height settings of the page sheet style.

Setting the Page Scale Width/Height
"""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.modify.page.sheet import ScalePagesWidthHeight, CalcStylePageKind
    
        # .. other code
        style = ScalePagesWidthHeight(width=2, height=3, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236646797-35f67919-24b7-4bb7-8f91-513d76f43e38:

    .. figure:: https://user-images.githubusercontent.com/4193389/236646797-35f67919-24b7-4bb7-8f91-513d76f43e38.png
        :alt: Calc dialog Page Style Sheet Scale style image modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Sheet Scale style image modified

Getting the page scale width/height from a style
""""""""""""""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # .. other code
        style_obj = ScalePagesWidthHeight.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
        assert style_obj.prop_style_name == str(CalcStylePageKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Number of Pages
^^^^^^^^^^^^^^^

The :py:class:`~ooodev.format.calc.modify.page.sheet.ScaleNumOfPages` class sets the scale number of pages settings of the page sheet style.

Setting the Page by Number of Pages
"""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        from ooodev.format.calc.modify.page.sheet import ScaleNumOfPages, CalcStylePageKind
    
        # .. other code
        style = ScaleNumOfPages(pages=3, style_name=CalcStylePageKind.DEFAULT)
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Style results.

.. cssclass:: screen_shot

    .. _236647040-ba3f5ee4-5dc6-4643-b749-b342a2592501:

    .. figure:: https://user-images.githubusercontent.com/4193389/236647040-ba3f5ee4-5dc6-4643-b749-b342a2592501.png
        :alt: Calc dialog Page Style Sheet Scale style image modified
        :figclass: align-center
        :width: 450px

        Calc dialog Page Style Sheet Scale style image modified

Getting the Page Sheet Scale Number of pages from a style
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # .. other code
        style_obj = ScaleNumOfPages.from_style(doc=doc, style_name=CalcStylePageKind.DEFAULT)
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
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.calc.modify.page.sheet.Order`
        - :py:class:`ooodev.format.calc.modify.page.sheet.Printing`
        - :py:class:`ooodev.format.calc.modify.page.sheet.ScaleReduceEnlarge`
        - :py:class:`ooodev.format.calc.modify.page.sheet.ScaleNumOfPages`
        - :py:class:`ooodev.format.calc.modify.page.sheet.ScalePagesWidthHeight`
        - :py:class:`ooodev.format.calc.modify.page.sheet.ScaleReduceEnlarge`
        - :py:class:`ooodev.format.calc.modify.page.sheet.ScalePagesWidthHeight`
        - :py:class:`ooodev.format.calc.modify.page.sheet.ScaleNumOfPages`