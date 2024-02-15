.. _help_chart2_format_direct_title_alignment:

Chart2 Direct Title/Subtitle Alignment
======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The ``style_orientation()`` and ``style_write_mode()`` methods are used to set the alignment of the chart title and subtitle.

Setup
-----

General setup used to run the examples in this page.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 40,41

        from __future__ import annotations
        from pathlib import Path
        import uno
        from ooo.dyn.awt.gradient_style import GradientStyle
        from ooodev.calc import CalcDoc, ZoomKind
        from ooodev.loader.lo import Lo
        from ooodev.utils.color import StandardColor
        from ooodev.utils.data_type.color_range import ColorRange
        from ooodev.format.chart2.direct.title.alignment import DirectionModeKind

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

                title.style_orientation(angle=15, vertical=False)
                title.style_write_mode(mode=DirectionModeKind.LR_TB)

                Lo.delay(1_000)
                doc.close()
            return 0

        if __name__ == "__main__":
            SystemExit(main())

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

Setting Direction and Orientation
---------------------------------

Before formatting the chart is seen in :numref:`686ff974-65de-4b94-8fc2-201206d048da`.

Apply to Title
^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.alignment import DirectionModeKind
        # ... other code
        title = chart_doc.get_title()
        if title is None:
            raise ValueError("Title not found")
        title.style_orientation(angle=15, vertical=False)
        title.style_write_mode(mode=DirectionModeKind.LR_TB)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`28f576a5-d385-492a-996e-995f66965dd3_1` and :numref:`e92ab05a-6093-43ce-a83b-14862827ec35_1`.

.. cssclass:: screen_shot

    .. _28f576a5-d385-492a-996e-995f66965dd3_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/28f576a5-d385-492a-996e-995f66965dd3
        :alt: Chart with title orientation set
        :figclass: align-center
        :width: 450px

        Chart with title orientation set

.. cssclass:: screen_shot

    .. _e92ab05a-6093-43ce-a83b-14862827ec35_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/e92ab05a-6093-43ce-a83b-14862827ec35
        :alt: Chart Title Alignment Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Alignment Dialog

Apply to Subtitle
^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        from ooodev.format.chart2.direct.title.alignment import DirectionModeKind
        # ... other code
        sub_title = chart_doc.first_diagram.get_title()
        if sub_title is None:
            raise ValueError("Title not found")
        sub_title.style_orientation(angle=15, vertical=False)
        sub_title.style_write_mode(mode=DirectionModeKind.LR_TB)

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`207076c0-ac22-4aef-a195-e5023ac04d64_1`.

.. cssclass:: screen_shot

    .. _207076c0-ac22-4aef-a195-e5023ac04d64_1:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/207076c0-ac22-4aef-a195-e5023ac04d64
        :alt: Chart with subtitle orientation set
        :figclass: align-center
        :width: 450px

        Chart with subtitle orientation set

Related Topics
--------------

.. seealso::

    .. cssclass:: ul-list

        - :ref:`part05`
        - :ref:`help_format_format_kinds`
        - :ref:`help_format_coding_style`
        - :ref:`help_chart2_format_direct_title`
        - :py:class:`~ooodev.loader.Lo`
        - :py:meth:`CalcSheet.dispatch_recalculate() <ooodev.calc.calc_sheet.CalcSheet.dispatch_recalculate>`