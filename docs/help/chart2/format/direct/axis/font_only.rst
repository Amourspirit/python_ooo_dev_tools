.. _help_chart2_format_direct_axis_font_only:

Chart2 Direct Axis Font Only
============================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_font()`` method gives similar options as :numref:`618f07b1-23a8-46b9-9aea-0ef4243bb0e2` Font Dialog, but without the dialog.

Setup
-----

General setup for this example.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 29, 30, 31

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
                _ = chart_doc.style_border_line(color=StandardColor.GREEN_DARK2, width=0.9)
                _ = chart_doc.style_area_gradient(
                    step_count=0,
                    offset=Offset(41, 50),
                    style=GradientStyle.RADIAL,
                    grad_color=ColorRange(StandardColor.TEAL, StandardColor.YELLOW_DARK1),
                )
                _ = chart_doc.axis_y.style_font(
                    name="Lucida Calligraphy", size=14, font_style="italic"
                )

                Lo.delay(1_000)
                doc.close()
            return 0


        if __name__ == "__main__":
            SystemExit(main())


    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None


Apply Font Only
---------------

Before formatting the chart is seen in :numref:`3adb4ebc-83d9-44c6-9bba-6c92e11f3b0a`.

Apply to Y-Axis
"""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        _ = chart_doc.axis_y.style_font(
            name="Lucida Calligraphy", size=14, font_style="italic"
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`4f6b0e7e-c772-4f57-9cf7-8971dc88c2a3_1` and :numref:`618f07b1-23a8-46b9-9aea-0ef4243bb0e2_1`.

.. cssclass:: screen_shot

    .. _4f6b0e7e-c772-4f57-9cf7-8971dc88c2a3_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/4f6b0e7e-c772-4f57-9cf7-8971dc88c2a3
        :alt: Chart with Y-Axis Font set
        :figclass: align-center
        :width: 450px

        Chart with Y-Axis Font set


.. cssclass:: screen_shot

    .. _618f07b1-23a8-46b9-9aea-0ef4243bb0e2_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/618f07b1-23a8-46b9-9aea-0ef4243bb0e2
        :alt: Chart Y-Axis Dialog Font
        :figclass: align-center
        :width: 450px

        Chart Y-Axis Dialog Font

Apply to X-Axis
"""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        _ = chart_doc.axis_x.style_font(
            name="Lucida Calligraphy", size=14, font_style="italic"
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Running the above code will produce the following output shown in :numref:`1090675b-1837-412e-b430-c0519a460c18_1`.

.. cssclass:: screen_shot

    .. _1090675b-1837-412e-b430-c0519a460c18_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/1090675b-1837-412e-b430-c0519a460c18
        :alt: Chart with Y-Axis Font set
        :figclass: align-center
        :width: 450px

        Chart with Y-Axis Font set

Getting the font style
""""""""""""""""""""""

For for all the Axis properties you can get the font style using the ``style_font_get()`` method.


.. tabs::

    .. code-tab:: python

        # ... other code
        f_style = chart_doc.axis_y.style_font_get()
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
        - :ref:`help_chart2_format_direct_axis_font_effects`
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_x_axis() <ooodev.office.chart2.Chart2.style_x_axis>`
        - :py:meth:`Chart2.style_x_axis2() <ooodev.office.chart2.Chart2.style_x_axis2>`
        - :py:meth:`Chart2.style_y_axis() <ooodev.office.chart2.Chart2.style_y_axis>`
        - :py:meth:`Chart2.style_y_axis2() <ooodev.office.chart2.Chart2.style_y_axis2>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.axis.font.FontOnly`