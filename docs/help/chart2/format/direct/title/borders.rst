.. _help_chart2_format_direct_title_borders:

Chart2 Direct Title/Subtitle Borders
====================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_border_line()`` method has the same options as the Chart Data Series Borders dialog
as seen in :numref:`a31ee22f-14cc-43ef-844f-7a078ec1abd9`.

Setup
-----

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 39,40,41,42

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor
        from ooodev.utils.data_type.color_range import ColorRange

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                fnm = Path.cwd() / "tmp" / "piechart.ods"
                doc = CalcDoc.open_doc(fnm=fnm, visible=True)
                Lo.delay(500)
                doc.zoom(ZoomKind.ZOOM_100_PERCENT)

                sheet = doc.sheets[0]
                sheet["A1"].goto()
                chart_table = sheet.charts[0]
                chart_doc = chart_table.chart_doc
                _ = chart_doc.style_border_line(
                    color=StandardColor.PURPLE_DARK1,
                    width=0.7,
                )
                _ = chart_doc.style_area_gradient(
                    step_count=64,
                    style=GradientStyle.SQUARE,
                    angle=45,
                    grad_color=ColorRange(
                        StandardColor.BLUE_DARK1,
                        StandardColor.PURPLE_LIGHT2,
                    ),
                )

                title = chart_doc.get_title()
                if title is None:
                    raise ValueError("Title not found")

                title.style_border_line(
                    color=StandardColor.MAGENTA_DARK1,
                    width=0.75,
                )

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Applying Line Properties
------------------------

The ``style_border_line()`` method is called to set the title and subtitle border line properties.

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

Apply to Title
""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code

        title = chart_doc.get_title()
        if title is None:
            raise ValueError("Title not found")

        title.style_border_line(
            color=StandardColor.MAGENTA_DARK1,
            width=0.75,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`9b8faf7e-9cfa-407d-880c-1efce5b012fe_1` and :numref:`a31ee22f-14cc-43ef-844f-7a078ec1abd9_1`.


.. cssclass:: screen_shot

    .. _9b8faf7e-9cfa-407d-880c-1efce5b012fe_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/9b8faf7e-9cfa-407d-880c-1efce5b012fe
        :alt: Chart with title border set
        :figclass: align-center
        :width: 450px

        Chart with title border set

.. cssclass:: screen_shot

    .. _a31ee22f-14cc-43ef-844f-7a078ec1abd9_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/a31ee22f-14cc-43ef-844f-7a078ec1abd9
        :alt: Chart Data Series Borders Default Dialog
        :figclass: align-center
        :width: 450px

        Chart Data Series Borders Default Dialog

Apply to Subtitle
"""""""""""""""""

.. tabs::

    .. code-tab:: python

        # ... other code
        sub_title = chart_doc.first_diagram.get_title()
        if sub_title is None:
            raise ValueError("Title not found")

        sub_title.style_border_line(
            color=StandardColor.MAGENTA_DARK1,
            width=0.75,
        )

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`27378b9f-41c0-4975-8b14-161133e81ca0_1`.


.. cssclass:: screen_shot

    .. _27378b9f-41c0-4975-8b14-161133e81ca0_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/27378b9f-41c0-4975-8b14-161133e81ca0
        :alt: Chart with subtitle border set
        :figclass: align-center
        :width: 450px

        Chart with subtitle border set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :py:class:`~ooodev.gui.GUI`
        - :py:class:`~ooodev.loader.Lo`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`
