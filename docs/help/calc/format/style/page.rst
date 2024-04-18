.. _help_calc_format_style_page:

Calc Style Page
===============

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2


Overview
--------

Applying Page Styles can be accomplished using the :py:class:`ooodev.format.calc.style.Page` class.

The :py:class:`~ooodev.format.calc.style.page.kind.calc_style_page_kind.CalcStylePageKind` enum is used to lookup the style to be applied.

Setup
-----

General function used to run these examples.

.. tabs::

    .. code-tab:: python

        import uno
        from ooodev.format.calc.style import Page, CalcStylePageKind
        from ooodev.office.calc import Calc
        from ooodev.gui import GUI
        from ooodev.loader.lo import Lo


        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 400)

                sheet = Calc.get_active_sheet()

                cell_obj = Calc.get_cell_obj("A1")
                Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)

                style = Page(name=CalcStylePageKind.REPORT)
                style.apply(sheet)

                page_style = Page.from_obj(sheet)
                assert page_style.prop_name == str(CalcStylePageKind.REPORT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply Style to a sheet
----------------------

.. tabs::

    .. code-tab:: python

        # ... other code
        style = Page(name=CalcStylePageKind.REPORT)
        style.apply(sheet)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Result be seen in :numref:`236843614-64366f7c-c4ae-44d0-8377-eab350e6a4a4`.

.. cssclass:: screen_shot

    .. _236843614-64366f7c-c4ae-44d0-8377-eab350e6a4a4:

    .. figure:: https://user-images.githubusercontent.com/4193389/236843614-64366f7c-c4ae-44d0-8377-eab350e6a4a4.png
        :alt: Style applied to Cell
        :figclass: align-center
        :width: 550px

        Style applied to Cell

Get Style from a Cell
---------------------

Get Style from a sheet by calling ``Page.from_obj()`` passing in the ``sheet`` object.

.. tabs::

    .. code-tab:: python

        # ... other code

        page_style = Page.from_obj(sheet)
        assert page_style.prop_name == str(CalcStylePageKind.REPORT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:class:`ooodev.format.calc.style.Page`
