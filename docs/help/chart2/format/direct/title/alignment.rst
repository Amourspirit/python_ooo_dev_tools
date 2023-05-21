.. _help_chart2_format_direct_title_alignment:

Chart2 Direct Title/Subtitle Alignment
======================================

.. contents:: Table of Contents
    :local:
    :backlinks: none
    :depth: 2

Overview
--------

The :py:class:`ooodev.format.chart2.direct.title.alignment.Direction` and :py:class:`ooodev.format.chart2.direct.title.alignment.Orientation`
classes are used to set the alignment of the chart title.

Calls to the :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>` and
:py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>` methods are used to set the Title and Subtitle formatting of a Chart.


Setup
-----

General setup used to run the examples in this page.

.. tabs::

    .. code-tab:: python
        :emphasize-lines: 34,35,36

        import uno
        from ooodev.format.chart2.direct.title.alignment import Direction, Orientation, DirectionModeKind
        from ooodev.format.chart2.direct.general.borders import LineProperties as ChartLineProperties
        from ooodev.format.chart2.direct.general.area import Gradient as ChartGradient
        from ooodev.format.chart2.direct.general.area import Gradient as GradientStyle, ColorRange
        from ooodev.office.calc import Calc
        from ooodev.office.chart2 import Chart2
        from ooodev.utils.color import StandardColor
        from ooodev.utils.gui import GUI
        from ooodev.utils.lo import Lo

        def main() -> int:
            with Lo.Loader(connector=Lo.ConnectPipe()):
                doc = Calc.open_doc("pie_flat_chart.ods")
                GUI.set_visible(True, doc)
                Lo.delay(500)
                Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

                sheet = Calc.get_active_sheet()

                Calc.goto_cell(cell_name="A1", doc=doc)
                chart_doc = Chart2.get_chart_doc(sheet=sheet, chart_name="pie_chart")

                chart_bdr_line = ChartLineProperties(color=StandardColor.PURPLE_DARK1, width=0.7)
                chart_grad = ChartGradient(
                    chart_doc=chart_doc,
                    step_count=64,
                    style=GradientStyle.SQUARE,
                    angle=45,
                    grad_color=ColorRange(StandardColor.BLUE_DARK1, StandardColor.PURPLE_LIGHT2),
                )
                Chart2.style_background(chart_doc=chart_doc, styles=[chart_grad, chart_bdr_line])

                title_orient = Orientation(angle=15, vertical=False)
                title_dir = Direction(mode=DirectionModeKind.LR_TB)
                Chart2.style_title(chart_doc=chart_doc, styles=[title_orient, title_dir])

                Lo.delay(1_000)
                Lo.close_doc(doc)
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

        import uno
        from ooodev.format.chart2.direct.title.alignment import Direction, Orientation, DirectionModeKind
        # ... other code

        title_orient = Orientation(angle=15, vertical=False)
        title_dir = Direction(mode=DirectionModeKind.LR_TB)
        Chart2.style_title(chart_doc=chart_doc, styles=[title_orient, title_dir])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`28f576a5-d385-492a-996e-995f66965dd3` and :numref:`e92ab05a-6093-43ce-a83b-14862827ec35`.

.. cssclass:: screen_shot

    .. _28f576a5-d385-492a-996e-995f66965dd3:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/28f576a5-d385-492a-996e-995f66965dd3
        :alt: Chart with title orientation set
        :figclass: align-center
        :width: 450px

        Chart with title orientation set

.. cssclass:: screen_shot

    .. _e92ab05a-6093-43ce-a83b-14862827ec35:

    .. figure:: https://github.com/Amourspirit/python_ooo_dev_tools/assets/4193389/e92ab05a-6093-43ce-a83b-14862827ec35
        :alt: Chart Title Alignment Dialog
        :figclass: align-center
        :width: 450px

        Chart Title Alignment Dialog

Apply to Subtitle
^^^^^^^^^^^^^^^^^

.. tabs::

    .. code-tab:: python

        # ... other code
        Chart2.style_subtitle(chart_doc=chart_doc, styles=[title_orient, title_dir])

    .. only:: html

        .. cssclass:: tab-none

            .. group-tab:: None

The results are seen in :numref:`207076c0-ac22-4aef-a195-e5023ac04d64`.

.. cssclass:: screen_shot

    .. _207076c0-ac22-4aef-a195-e5023ac04d64:

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
        - :py:class:`~ooodev.utils.gui.GUI`
        - :py:class:`~ooodev.utils.lo.Lo`
        - :py:class:`~ooodev.office.chart2.Chart2`
        - :py:meth:`Chart2.style_background() <ooodev.office.chart2.Chart2.style_background>`
        - :py:meth:`Chart2.style_title() <ooodev.office.chart2.Chart2.style_title>`
        - :py:meth:`Chart2.style_subtitle() <ooodev.office.chart2.Chart2.style_subtitle>`
        - :py:meth:`Calc.dispatch_recalculate() <ooodev.office.calc.Calc.dispatch_recalculate>`
        - :py:class:`ooodev.format.chart2.direct.title.alignment.Orientation`
        - :py:class:`ooodev.format.chart2.direct.title.alignment.Direction`