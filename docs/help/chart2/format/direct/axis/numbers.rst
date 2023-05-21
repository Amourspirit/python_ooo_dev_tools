.. _help_chart2_format_direct_axis_numbers:

Chart2 Direct Axis Numbers
==========================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.axis.numbers.Numbers` class is used to format the number of an axis.

Methods for formatting the number of an axis are:

    - :py:meth:`Chart2.style_x_axis() <ooodev.office.chart2.Chart2.style_x_axis>`
    - :py:meth:`Chart2.style_x_axis2() <ooodev.office.chart2.Chart2.style_x_axis2>`
    - :py:meth:`Chart2.style_y_axis() <ooodev.office.chart2.Chart2.style_y_axis>`
    - :py:meth:`Chart2.style_y_axis2() <ooodev.office.chart2.Chart2.style_y_axis2>`


Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 34,35,36,37

        import uno
        from ooodev.format.chart2.direct.axis.numbers import Numbers, NumberFormatIndexEnum
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient
        from ooodev.format.chart2.direct.general.area import GradientStyle, ColorRange, Offset
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc(Path.cwd() / "tmp" / "bon_voyage.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="Object 1")

                chart_bdr_line = ChartLineProperties(color=StandardColor.GREEN_DARK2, width=0.9)
                chart_grad = ChartGradient(
                    chart_doc=chart_doc,
                    step_count=0,
                    offset=Offset(41, 50),
                    style=GradientStyle.RADIAL,
                    grad_color=ColorRange(StandardColor.TEAL, StandardColor.YELLOW_DARK1),
                )
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                num_style = Numbers(
                    chart_doc, source_format=False, num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2
                )
                Chart2.style_y_axis(chart_doc=chart_doc, styles=[num_style])

                Lo.delay(1_000)
                Lo.close_doc(doc)
            return 0

        if __name__ == "__main__":
            SystemExit(main())
    
    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply to Axis
-------------

Before formatting the chart is seen in :numref:`3adb4ebc-83d9-44c6-9bba-6c92e11f3b0a`.

Apply to Y-Axis
"""""""""""""""

The ``NumberFormatIndexEnum`` enum contains the values in |num_fmt_index|_ for easy lookup.

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.axis.numbers import Numbers, NumberFormatIndexEnum
        # .. other code

        num_style = Numbers(
            chart_doc, source_format=False, num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2
        )
        Chart2.style_y_axis(chart_doc=chart_doc, styles=[num_style])

    
    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`602db3dc-9afd-4a9a-860c-d8bc4c75e5da` and :numref:`4f2d29a6-3320-40fb-ae3d-a397c8c27998`.


.. cssclass:: screen_shot

    .. _602db3dc-9afd-4a9a-860c-d8bc4c75e5da:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/602db3dc-9afd-4a9a-860c-d8bc4c75e5da
        :alt: Chart with Y-Axis Formatted to Currency with two decimal places
        :figclass: align-center
        :width: 450px

        Chart with Y-Axis Formatted to Currency with two decimal places

.. cssclass:: screen_shot

    .. _4f2d29a6-3320-40fb-ae3d-a397c8c27998:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4f2d29a6-3320-40fb-ae3d-a397c8c27998
        :alt: Chart Area Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Area Borders Default Dialog

Apply to Secondary Y-Axis
"""""""""""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_y_axis2(chart_doc=chart_doc, styles=[num_style])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`d572bc21-c52a-4d94-8e79-72b373b56060`.


.. cssclass:: screen_shot

    .. _d572bc21-c52a-4d94-8e79-72b373b56060:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/d572bc21-c52a-4d94-8e79-72b373b56060
        :alt: Chart with Y-Axis Formatted to Currency with two decimal places
        :figclass: align-center
        :width: 450px

        Chart with Y-Axis Formatted to Currency with two decimal places

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_axis`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_x_axis() <ooodev.office.chart2.Chart2.style_x_axis>`
        - :py:meth:`Chart2.style_x_axis2() <ooodev.office.chart2.Chart2.style_x_axis2>`
        - :py:meth:`Chart2.style_y_axis() <ooodev.office.chart2.Chart2.style_y_axis>`
        - :py:meth:`Chart2.style_y_axis2() <ooodev.office.chart2.Chart2.style_y_axis2>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.axis.numbers.Numbers`

.. |num_fmt| replace:: API NumberFormat
.. _num_fmt: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1NumberFormat.html

.. |num_fmt_index| replace:: API NumberFormatIndex
.. _num_fmt_index: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1i18n_1_1NumberFormatIndex.html
