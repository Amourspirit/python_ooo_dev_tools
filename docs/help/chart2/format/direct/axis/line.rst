.. _help_chart2_format_direct_axis_line:

Chart2 Direct Axis Line
=======================


.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.axis.line.LineProperties` class gives the same options
as seen in :numref:`ee039628-896d-4ab2-a4c8-843d643a5579`.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 35,36,37

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.awt.gradient_style import GradientStyle
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
                _ = chart_doc.axis_y.style_axis_line(
                    color=StandardColor.TEAL, width=0.75
                )

                Lo.delay(1_000)
                doc.close()
            return 0


        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Apply Line to Axis
------------------

The ``style_axis_line()`` method is used to set the Axis line properties.

Before formatting the chart is seen in :numref:`3adb4ebc-83d9-44c6-9bba-6c92e11f3b0a`.

Apply to Y-Axis
"""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        _ = chart_doc.axis_y.style_axis_line(
            color=StandardColor.TEAL, width=0.75
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`1c1711ce-1169-4106-8925-c7790dbad0e0_1` and :numref:`ee039628-896d-4ab2-a4c8-843d643a5579_1`


.. cssclass:: screen_shot

    .. _1c1711ce-1169-4106-8925-c7790dbad0e0_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1c1711ce-1169-4106-8925-c7790dbad0e0
        :alt: Chart with Y-Axis line set
        :figclass: align-center
        :width: 450px

        Chart with Y-Axis line set

.. cssclass:: screen_shot

    .. _ee039628-896d-4ab2-a4c8-843d643a5579_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ee039628-896d-4ab2-a4c8-843d643a5579
        :alt: Chart Y-Axis Line Dialog
        :figclass: align-center
        :width: 450px

        Chart Y-Axis Line Dialog

Apply to X-Axis
"""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        # ... other code
        _ = chart_doc.axis_x.style_axis_line(
            color=StandardColor.TEAL, width=0.75
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`ae063a29-1ebc-442e-9f9b-7d9dba8f64ad_1`


.. cssclass:: screen_shot

    .. _ae063a29-1ebc-442e-9f9b-7d9dba8f64ad_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/ae063a29-1ebc-442e-9f9b-7d9dba8f64ad
        :alt: Chart with Y-Axis line set
        :figclass: align-center
        :width: 450px

        Chart with Y-Axis line set

Getting the Axis Line style
"""""""""""""""""""""""""""

For for all the Axis properties you can get the line style using the ``style_axis_line()`` method.


.. tabs::

    .. code-tab:: python

        # ... other code
        f_style = chart_doc.axis_y.style_axis_line()
        assert f_style is not None

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_axis`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.axis.line.LineProperties`