.. _help_chart2_format_direct_axis_numbers:

Chart2 Direct Axis Numbers
==========================

.. contents:: Table of Contents
    :local:
    :backlinks: top
    :depth: 2

Overview
--------

The ``style_numbers_numbers()`` method is used to format the number of an axis.

Setup
-----

General setup for examples.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 36,37,38,39

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.utils.color import StandardColor
        from ooodev.loader.lo import Lo
        from ooodev.utils.data_type.color_range import ColorRange
        from ooodev.utils.data_type.offset import Offset

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "bon_voyage.ods"
                doc = CalcDoc.open_doc(fnm=fnm, visible=True)
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)

                sheet = doc.sheets[0]
                sheet["A1"].goto()
                chart_table = sheet.charts[0]
                chart_doc = chart_table.chart_doc
                _ = chart_doc.style_border_line(
                    color=StandardColor.GREEN_DARK2,
                    width=0.9,
                )
                _ = chart_doc.style_area_gradient(
                    step_count=0,
                    offset=Offset(41, 50),
                    style=GradientStyle.RADIAL,
                    grad_color=ColorRange(
                        StandardColor.TEAL,
                        StandardColor.YELLOW_DARK1,
                    ),
                )
                _ = chart_doc.axis_y.style_numbers_numbers(
                    source_format=False,
                    num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2,
                )

                Lo.delay(1_000)
                doc.close()
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

        from ooo.dyn.i18n.number_format_index import NumberFormatIndexEnum
        # .. other code

        _ = chart_doc.axis_y.style_numbers_numbers(
            source_format=False,
            num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2,
        )

    
    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`602db3dc-9afd-4a9a-860c-d8bc4c75e5da_1` and :numref:`4f2d29a6-3320-40fb-ae3d-a397c8c27998_1`.


.. cssclass:: screen_shot

    .. _602db3dc-9afd-4a9a-860c-d8bc4c75e5da_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/602db3dc-9afd-4a9a-860c-d8bc4c75e5da
        :alt: Chart with Y-Axis Formatted to Currency with two decimal places
        :figclass: align-center
        :width: 450px

        Chart with Y-Axis Formatted to Currency with two decimal places

.. cssclass:: screen_shot

    .. _4f2d29a6-3320-40fb-ae3d-a397c8c27998_1:

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
        y2_axis = chart_doc.axis2_y
        if y2_axis is not None:
            _ = y2_axis.style_numbers_numbers(
                source_format=False,
                num_format_index=NumberFormatIndexEnum.CURRENCY_1000DEC2,
            )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`d572bc21-c52a-4d94-8e79-72b373b56060_1`.


.. cssclass:: screen_shot

    .. _d572bc21-c52a-4d94-8e79-72b373b56060_1:

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
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`

.. |num_fmt| replace:: API NumberFormat
.. _num_fmt: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1util_1_1NumberFormat.html

.. |num_fmt_index| replace:: API NumberFormatIndex
.. _num_fmt_index: https://api.libreoffice.org/docs/idl/ref/namespacecom_1_1sun_1_1star_1_1i18n_1_1NumberFormatIndex.html
