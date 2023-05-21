.. _help_calc_format_modify_cell_numbers:

Calc Modify Cell Numbers
========================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.calc.modify.cell.numbers.Numbers` class is used to programmatically set the number properties of a style
as seen dialog  shown in :numref:`236624645-95bd9940-7e36-4ee4-8a6d-1d762e876dd8`.

.. cssclass:: screen_shot

    .. _236624645-95bd9940-7e36-4ee4-8a6d-1d762e876dd8:

    .. figure:: https://user-images.githubusercontent.com/4193389/236624645-95bd9940-7e36-4ee4-8a6d-1d762e876dd8.png
        :alt: Calc Format Cell dialog Style Cell Numbers
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Style Cell Numbers

For more examples of setting number properties see :ref:`help_calc_format_direct_cell_numbers`.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 15, 16, 17, 18, 19

        import uno
        from ooodev.office.calc import Calc
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo
        from ooodev.format.calc.modify.cell.numbers import Numbers
        from ooodev.format.calc.modify.cell.numbers import StyleCellKind, NumberFormatIndexEnum

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectSocket()):
                doc = Calc.create_doc()
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom_value(doc, 100)

                style = Numbers(
                    num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2_RED,
                    style_name=StyleCellKind.DEFAULT,
                )
                style.apply(doc)

                style_obj = Numbers.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
                assert style_obj.prop_style_name == str(StyleCellKind.DEFAULT)

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting the style Number properties
-----------------------------------

.. tabs::

    .. code-tab:: python

        style = Numbers(
            num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2_RED,
            style_name=StyleCellKind.DEFAULT,
        )
        style.apply(doc)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output in :numref:`236624866-b89f0695-63d0-4d3a-8763-9f6f48cabe66`.

.. cssclass:: screen_shot

    .. _236624866-b89f0695-63d0-4d3a-8763-9f6f48cabe66:

    .. figure:: https://user-images.githubusercontent.com/4193389/236624866-b89f0695-63d0-4d3a-8763-9f6f48cabe66.png
        :alt: Calc Format Cell dialog Style Cell Numbers set
        :figclass: align-center
        :width: 450px

        Calc Format Cell dialog Style Cell Numbers set

Getting cell Number properties from a style
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code

        style_obj = Numbers.from_style(doc=doc, style_name=StyleCellKind.DEFAULT)
        assert style_obj.prop_style_name == str(StyleCellKind.DEFAULT)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - `API NumberFormat <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1NumberFormat.html>`__
        - `API NumberFormatIndex <https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1i18n_1_1NumberFormatIndex.html>`__
        - :ref:`help_calc_format_direct_cell_numbers`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`ooodev.format.calc.modify.cell.numbers.Numbers`
